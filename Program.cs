using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace PDFTronEML2DOC
{
    class Program
    {
        /// <summary>
        /// Get the starting index of the pattern
        /// </summary>
        /// <param name="pattern">byte pattern to find start of</param>
        /// <param name="buf">buffer to search</param>
        /// <param name="offsetInBuf">starting location in buffer</param>
        /// <returns>offset of start of pattern, or -1 if not found</returns>
        static int IndexOf(byte[] pattern, byte[] buf, int offsetInBuf)
        {
            if (pattern.Length == 0) return -1;
            if (buf.Length == 0) return -1;
            if (offsetInBuf + pattern.Length > buf.Length) return -1;
            int patternOffset = 0;
            int ptr = offsetInBuf;
            while (ptr < buf.Length)
            {
                if (buf[ptr] == pattern[patternOffset])
                {
                    if (patternOffset == pattern.Length - 1) return ptr - patternOffset;
                    ++patternOffset;
                }
                else
                {
                    patternOffset = 0;
                }
                ++ptr;
            }
            return -1;
        }

        /// <summary>
        /// Check if byte is an end of line byte
        /// </summary>
        /// <param name="b">Byte to check</param>
        /// <returns>true if end of line byte</returns>
        static bool IsEOL(byte b)
        {
            return b == 0x0A || b == 0x0D;
        }

        /// <summary>
        /// Find first byte that is an end of line byte
        /// </summary>
        /// <param name="buf">buffer to search</param>
        /// <param name="offsetInBuf">offset in buffer to start search at</param>
        /// <returns>first end of line byte after offsetInBuf, or -1 if end of buffer hit</returns>
        static int IndexOfEOL(byte[] buf, int offsetInBuf)
        {
            int offset = offsetInBuf;
            while (offset < buf.Length)
            {
                if (IsEOL(buf[offset])) return offset;
                ++offset;
            }
            return -1;
        }

        /// <summary>
        /// Checks if byte is Linear Whitespace (though skipping EOL bytes)
        /// </summary>
        /// <param name="buf">buffer to search</param>
        /// <param name="offsetInBuf">starting location in buffer</param>
        /// <returns>true if LWSP byte</returns>
        static bool IsLWSP(byte[] buf, int offsetInBuf)
        {
            byte b = buf[offsetInBuf];
            return b == 0x20 || b == 0x09;
        }

        /// <summary>
        /// Convert EML file to DOC using Outlook interop
        /// </summary>
        /// <param name="emlFile">Path to EML file to convert</param>
        /// <param name="tempPath">Path that function can use as temporary write/read location</param>
        /// <param name="docPath">Path to write DOC file to</param>
        /// <returns>True on creation of DOC file, otherwise false</returns>
        static bool EmlToDoc(string emlFile, string tempPath, string docPath)
        {
            bool success = false;

            var guid = Guid.NewGuid();
            string guidStr = guid.ToString();
            byte[] guidBytes = Encoding.ASCII.GetBytes(guidStr);

            byte[] emlData = System.IO.File.ReadAllBytes(emlFile);

            byte[] returnPathPattern = new byte[] { 0x52, 0x65, 0x74, 0x75, 0x72, 0x6E, 0x2D, 0x50, 0x61, 0x74, 0x68, 0x3A };

            byte[] subjectPattern = new byte[] { 0x0D, 0x0A, 0x53, 0x75, 0x62, 0x6A, 0x65, 0x63, 0x74, 0x3A, 0x20 };
            byte[] nextLinePattern = new byte[] { 0x0D, 0x0A };

            List<byte> subjectBuffer = new List<byte>();

            // TODO does not handle where Subject: is first line of file
            string originalSubject = "";
            int lastPos = -1;
            int indexOfSubject = IndexOf(subjectPattern, emlData, 0);
            if(indexOfSubject >= 0)
            {
                lastPos = indexOfSubject + subjectPattern.Length;
                int nextLine = IndexOf(nextLinePattern, emlData, lastPos);
                while (nextLine > 0)
                {
                    for (int i = lastPos; i < nextLine; ++i)
                    {
                        subjectBuffer.Add(emlData[i]);
                    }
                    lastPos = nextLine + 2;
                    if (!IsLWSP(emlData, lastPos)) break;

                    nextLine = IndexOf(nextLinePattern, emlData, lastPos);
                }
                originalSubject = System.Text.Encoding.UTF8.GetString(subjectBuffer.ToArray());
            }
            using (var stream = new FileStream(tempPath, FileMode.Create))
            {
                if(indexOfSubject >= 0)
                {
                    // write data from start to end of Subject header
                    stream.Write(emlData, 0, indexOfSubject + subjectPattern.Length);
                }
                else
                {
                    // inject subject to begining of stream
                    stream.Write(subjectPattern, 2, subjectPattern.Length - 2);
                    lastPos = 0;
                }
                // write guide
                stream.Write(guidBytes, 0, guidBytes.Length);
                stream.WriteByte(0x0D);
                stream.WriteByte(0x0A);
                // write remaining data
                stream.Write(emlData, lastPos, emlData.Length - lastPos);
            }

            ProcessStartInfo psi = new ProcessStartInfo(tempPath);
            psi.UseShellExecute = true;
            psi.CreateNoWindow = true;
            Process p = Process.Start(psi);
            //Thread.Sleep(100);

            Outlook.Application outlook = new Outlook.Application();
            Outlook.MailItem mailItem = null;
            int maxAttemptsToFindItem = 100;
            do
            {
                try
                {
                    Outlook.Inspectors inspectors = outlook.Inspectors;
                    foreach (Outlook.Inspector inspector in inspectors)
                    {
                        if (inspector == null) continue;
                        mailItem = (Outlook.MailItem)inspector.CurrentItem;
                        if (mailItem == null) continue;
                        string mailItemSubject = mailItem.Subject;
                        if (mailItemSubject != null && mailItemSubject.CompareTo(guidStr) == 0)
                        {
                            mailItem.Subject = originalSubject;
                            mailItem.SaveAs(docPath, Outlook.OlSaveAsType.olDoc);
                            success = true;
                            mailItem.Close(Outlook.OlInspectorClose.olDiscard);
                        }
                    }
                }
                catch (System.Exception e)
                {
                    Console.WriteLine("{0}\n{1}", emlFile, e);
                }
                --maxAttemptsToFindItem;
            } while (!success && maxAttemptsToFindItem > 0);
            if (!success && maxAttemptsToFindItem == 0) Console.WriteLine("{0} max attempts exceeded", emlFile);
            return success;
        }

        static void Main(string[] args)
        {
            if(args.Length < 2)
            {
                Console.WriteLine("usage: emlFileInput docFileOutput");
            }
            string tempPath = System.IO.Path.GetTempPath();
            var guid = Guid.NewGuid();
            string guidStr = guid.ToString();
            string tempFile = System.IO.Path.Combine(tempPath, guidStr + ".eml");
            try
            {
                bool success = EmlToDoc(args[0], tempFile, args[1]);
                if (success) Console.Write("[PASS] ");
                else Console.Write("[FAIL] ");
                Console.WriteLine("Converting {0} to {1}", args[0], args[1]);
            }
            catch (System.Exception e)
            {
                Console.WriteLine("Error converting {0}\n{1}", args[0], e);
            }
            finally
            {
                System.IO.File.Delete(tempFile);
            }
        }
    }
}
