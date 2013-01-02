/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
using System.IO;
using System;
using System.Collections.Generic;
namespace NPOI.XSSF.Util
{





    /**
     * This is a seriously sick fix for the fact that some .xlsx
     *  files contain raw bits of HTML, without being escaped
     *  or properly turned into XML.
     * The result is that they contain things like &gt;br&lt;,
     *  which breaks the XML parsing.
     * This very sick InputStream wrapper attempts to spot
     *  these go past, and fix them.
     * Only works for UTF-8 and US-ASCII based streams!
     * It should only be used where experience Shows the problem
     *  can occur...
     */
    public class EvilUnclosedBRFixingInputStream : Stream
    {
        private Stream source;
        private byte[] spare;

        private static byte[] detect = new byte[] {
              (byte)'<', (byte)'b', (byte)'r', (byte)'>'
           };

        public EvilUnclosedBRFixingInputStream(Stream source)
        {
            this.source = source;
        }

        /**
         * Warning - doesn't fix!
         */

        public int Read()
        {
            return source.ReadByte();
        }


        public override int Read(byte[] b, int off, int len)
        {
            // Grab any data left from last time
            int readA = ReadFromSpare(b, off, len);

            // Now read from the stream 
            int readB = source.Read(b, off + readA, len - readA);

            // Figure out how much we've done
            int read;
            if (readB == -1 || readB == 0)
            {
                read = readA;
            }
            else
            {
                read = readA + readB;
            }

            // Fix up our data
            if (read > 0)
            {
                read = fixUp(b, off, read);
            }

            // All done
            return read;
        }


        public int Read(byte[] b)
        {
            return this.Read(b, 0, b.Length);
        }

        /**
         * Reads into the buffer from the spare bytes
         */
        private int ReadFromSpare(byte[] b, int offset, int len)
        {
            if (spare == null) return 0;
            if (len == 0) throw new ArgumentException("Asked to read 0 bytes");

            if (spare.Length <= len)
            {
                // All fits, good
                Array.Copy(spare, 0, b, offset, spare.Length);
                int read = spare.Length;
                spare = null;
                return read;
            }
            else
            {
                // We have more spare than they can copy with...
                byte[] newspare = new byte[spare.Length - len];
                Array.Copy(spare, 0, b, offset, len);
                Array.Copy(spare, len, newspare, 0, newspare.Length);
                spare = newspare;
                return len;
            }
        }
        private void AddToSpare(byte[] b, int offset, int len, bool atTheEnd)
        {
            if (spare == null)
            {
                spare = new byte[len];
                Array.Copy(b, offset, spare, 0, len);
            }
            else
            {
                byte[] newspare = new byte[spare.Length + len];
                if (atTheEnd)
                {
                    Array.Copy(spare, 0, newspare, 0, spare.Length);
                    Array.Copy(b, offset, newspare, spare.Length, len);
                }
                else
                {
                    Array.Copy(b, offset, newspare, 0, len);
                    Array.Copy(spare, 0, newspare, len, spare.Length);
                }
                spare = newspare;
            }
        }

        private int fixUp(byte[] b, int offset, int read)
        {
            // Do we have any potential overhanging ones?
            for (int i = 0; i < detect.Length - 1; i++)
            {
                int base1 = offset + read - 1 - i;
                if (base1 < 0) continue;

                bool going = true;
                for (int j = 0; j <= i && going; j++)
                {
                    if (b[base1 + j] == detect[j])
                    {
                        // Matches
                    }
                    else
                    {
                        going = false;
                    }
                }
                if (going)
                {
                    // There could be a <br> handing over the end, eg <br|
                    AddToSpare(b, base1, i + 1, true);
                    read -= 1;
                    read -= i;
                    break;
                }
            }

            // Find places to fix
            List<int> fixAt = new List<int>();
            for (int i = offset; i <= offset + read - detect.Length; i++)
            {
                bool going = true;
                for (int j = 0; j < detect.Length && going; j++)
                {
                    if (b[i + j] != detect[j])
                    {
                        going = false;
                    }
                }
                if (going)
                {
                    fixAt.Add(i);
                }
            }

            if (fixAt.Count == 0)
            {
                return read;
            }

            // If there isn't space in the buffer to contain
            //  all the fixes, then save the overshoot for next time
            int needed = offset + read + fixAt.Count;
            int overshoot = needed - b.Length;
            if (overshoot > 0)
            {
                // Make sure we don't loose part of a <br>!
                int fixes = 0;
                foreach (int at in fixAt)
                {
                    if (at > offset + read - detect.Length - overshoot - fixes)
                    {
                        overshoot = needed - at - 1 - fixes;
                        break;
                    }
                    fixes++;
                }

                AddToSpare(b, offset + read - overshoot, overshoot, false);
                read -= overshoot;
            }

            // Fix them, in reverse order so the
            //  positions are valid
            for (int j = fixAt.Count - 1; j >= 0; j--)
            {
                int i = fixAt[j];
                if (i >= read + offset)
                {
                    // This one has Moved into the overshoot
                    continue;
                }
                if (i > read - 3)
                {
                    // This one has Moved into the overshoot
                    continue;
                }

                byte[] tmp = new byte[read - i - 3];
                Array.Copy(b, i + 3, tmp, 0, tmp.Length);
                b[i + 3] = (byte)'/';
                Array.Copy(tmp, 0, b, i + 4, tmp.Length);
                // It got one longer
                read++;
            }
            return read;
        }

        public override bool CanRead
        {
            get { return true; }
        }

        public override bool CanSeek
        {
            get { return true; }
        }

        public override bool CanWrite
        {
            get { return false; }
        }

        public override void Flush()
        {
            
        }

        public override long Length
        {
            get { return source.Length; }
        }

        public override long Position
        {
            get
            {
                return source.Position;
            }
            set
            {
                source.Position = value;
            }
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return source.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            throw new InvalidOperationException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new InvalidOperationException();
        }
    }
}


