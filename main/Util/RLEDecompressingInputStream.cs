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

namespace NPOI.Util
{
    using System;
    using System.IO;

    /**
     * Wrapper of InputStream which provides Run Length Encoding (RLE) 
     *  decompression on the fly. Uses MS-OVBA decompression algorithm. See
     * http://download.microsoft.com/download/2/4/8/24862317-78F0-4C4B-B355-C7B2C1D997DB/[MS-OVBA].pdf
     */
    public class RLEDecompressingInputStream : InputStream
    {

        /**
         * Bitmasks for performance
         */
        private static int[] POWER2 = new int[] {
            0x0001, // 2^0
            0x0002, // 2^1
            0x0004, // 2^2
            0x0008, // 2^3
            0x0010, // 2^4
            0x0020, // 2^5
            0x0040, // 2^6
            0x0080, // 2^7
            0x0100, // 2^8
            0x0200, // 2^9
            0x0400, // 2^10
            0x0800, // 2^11
            0x1000, // 2^12
            0x2000, // 2^13
            0x4000, // 2^14
            0x8000  // 2^15
    };

        /** the wrapped inputstream */
        private Stream input;

        /** a byte buffer with size 4096 for storing a single chunk */
        private byte[] buf;

        /** the current position in the byte buffer for Reading */
        private int pos;

        /** the number of bytes in the byte buffer */
        private int len;

        public override bool CanRead 
        {
            get {return input.CanRead; }
        }
        public override bool CanSeek
        {
            get {return input.CanSeek; }
        }
        public override bool CanWrite
        {
            get {return input.CanWrite; }
        }
        public override long Length
        {
            get {return input.Length; }
        }
        public override long Position
        {
            get { return input.Position; }
            set { input.Position = value;}
        }

        /**
         * Creates a new wrapper RLE Decompression InputStream.
         * 
         * @param in The stream to wrap with the RLE Decompression
         * @throws IOException
         */
        public RLEDecompressingInputStream(Stream input)
        {
            this.input = input;
            buf = new byte[4096];
            pos = 0;
            int header = input.ReadByte();
            if (header != 0x01)
            {
                throw new ArgumentException(String.Format("Header byte 0x01 expected, received 0x{0:X2}", header & 0xFF));
            }
            len = ReadChunk();
        }


        public override int Read()
        {
            if (len == -1)
            {
                return -1;
            }
            if (pos >= len)
            {
                if ((len = ReadChunk()) == -1)
                {
                    return -1;
                }
            }
            return buf[pos++];
        }


        public override int Read(byte[] b)
        {
            return Read(b, 0, b.Length);
        }


        public override int Read(byte[] b, int off, int l)
        {
            if (len == -1)
            {
                return -1;
            }
            int offset = off;
            int length = l;
            while (length > 0)
            {
                if (pos >= len)
                {
                    if ((len = ReadChunk()) == -1)
                    {
                        return offset > off ? offset - off : -1;
                    }
                }
                int c = Math.Min(length, len - pos);
                Array.Copy(buf, pos, b, offset, c);
                pos += c;
                length -= c;
                offset += c;
            }
            return l;
        }


        public override long Skip(long n)
        {
            long length = n;
            while (length > 0)
            {
                if (pos >= len)
                {
                    if ((len = ReadChunk()) == -1)
                    {
                        return -1;
                    }
                }
                int c = (int)Math.Min(n, len - pos);
                pos += c;
                length -= c;
            }
            return n;
        }


        public override int Available()
        {
            return (len > 0 ? len - pos : 0);
        }


        public override void Close()
        {
            input.Close();
        }

        /**
         * Reads a single chunk from the underlying inputstream.
         * 
         * @return number of bytes that were Read, or -1 if the end of the stream was reached.
         * @throws IOException
         */
        private int ReadChunk()
        {
            pos = 0;
            int w = ReadShort(input);
            if (w == -1)
            {
                return -1;
            }
            int chunkSize = (w & 0x0FFF) + 1; // plus 3 bytes minus 2 for the length
            if ((w & 0x7000) != 0x3000)
            {
                throw new ArgumentException(String.Format("Chunksize header A should be 0x3000, received 0x{0:X4}", w & 0xE000));
            }
            bool rawChunk = (w & 0x8000) == 0;
            if (rawChunk)
            {
                if (input.Read(buf, 0, chunkSize) < chunkSize)
                {
                    throw new InvalidOperationException(String.Format("Not enough bytes Read, expected {0}", chunkSize));
                }
                return chunkSize;
            }
            else
            {
                int inOffset = 0;
                int outOffset = 0;
                while (inOffset < chunkSize)
                {
                    int tokenFlags = input.ReadByte();
                    inOffset++;
                    if (tokenFlags == -1)
                    {
                        break;
                    }
                    for (int n = 0; n < 8; n++)
                    {
                        if (inOffset >= chunkSize)
                        {
                            break;
                        }
                        if ((tokenFlags & POWER2[n]) == 0)
                        {
                            // literal
                            int b = input.ReadByte();
                            if (b == -1)
                            {
                                return -1;
                            }
                            buf[outOffset++] = (byte)b;
                            inOffset++;
                        }
                        else
                        {
                            // compressed token
                            int token = ReadShort(input);
                            if (token == -1)
                            {
                                return -1;
                            }
                            inOffset += 2;
                            int copyLenBits = GetCopyLenBits(outOffset - 1);
                            int copyOffset = (token >> (copyLenBits)) + 1;
                            int copyLen = (token & (POWER2[copyLenBits] - 1)) + 3;
                            int startPos = outOffset - copyOffset;
                            int endPos = startPos + copyLen;
                            for (int i = startPos; i < endPos; i++)
                            {
                                buf[outOffset++] = buf[i];
                            }
                        }
                    }
                }
                return outOffset;
            }
        }

        /**
         * Helper method to determine how many bits in the CopyToken are used for the CopyLength.
         * 
         * @param offset
         * @return returns the number of bits in the copy token (a value between 4 and 12)
         */
        static int GetCopyLenBits(int offset)
        {
            for (int n = 11; n >= 4; n--)
            {
                if ((offset & POWER2[n]) != 0)
                {
                    return 15 - n;
                }
            }
            return 12;
        }

        /**
         * Convenience method for read a 2-bytes short in little endian encoding.
         * 
         * @return short value from the stream, -1 if end of stream is reached
         * @throws IOException
         */
        public int ReadShort()
        {
            return ReadShort(this);
        }

        /**
         * Convenience method for read a 4-bytes int in little endian encoding.
         * 
         * @return integer value from the stream, -1 if end of stream is reached
         * @throws IOException
         */
        public int ReadInt()
        {
            return ReadInt(this);
        }

        private int ReadShort(Stream stream)
        {
            int b0, b1;
            if ((b0 = stream.ReadByte()) == -1)
            {
                return -1;
            }
            if ((b1 = stream.ReadByte()) == -1)
            {
                return -1;
            }
            return (b0 & 0xFF) | ((b1 & 0xFF) << 8);
        }

        private int ReadInt(InputStream stream)
        {
            int b0, b1, b2, b3;
            if ((b0 = stream.Read()) == -1)
            {
                return -1;
            }
            if ((b1 = stream.Read()) == -1)
            {
                return -1;
            }
            if ((b2 = stream.Read()) == -1)
            {
                return -1;
            }
            if ((b3 = stream.Read()) == -1)
            {
                return -1;
            }
            return (b0 & 0xFF) | ((b1 & 0xFF) << 8) | ((b2 & 0xFF) << 16) | ((b3 & 0xFF) << 24);
        }

        public static byte[] Decompress(byte[] compressed)
        {
            return Decompress(compressed, 0, compressed.Length);
        }

        public static byte[] Decompress(byte[] compressed, int offset, int length)
        {
            MemoryStream out1 = new MemoryStream();
            Stream instream = new MemoryStream(compressed, offset, length);
            InputStream stream = new RLEDecompressingInputStream(instream);
            IOUtils.Copy(stream, out1);
            stream.Close();
            out1.Close();
            return out1.ToArray();
        }


        public override void Flush()
        {
            input.Flush();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return input.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            input.SetLength(value);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            input.Write(buffer, offset, count);
        }
    }

}