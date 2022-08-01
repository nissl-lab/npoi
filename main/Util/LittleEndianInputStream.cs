/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

    /// <summary>
    /// <para>Wraps an <see cref="T:System.IO.Stream"/> providing <see cref="T:NPOI.Util.ILittleEndianInput"/><p/></para>
    /// <para>
    /// This class does not buffer any input, so the stream Read position maintained
    /// by this class is consistent with that of the inner stream.
    /// </para>
    /// </summary>
    /// <remarks>
    /// @author Josh Micich
    /// </remarks>
    public class LittleEndianInputStream : FilterInputStream, ILittleEndianInput
    {
        private int readLimit = -1;
        private long markPos = -1;

        public LittleEndianInputStream(Stream is1) : base(new FileInputStream(is1))
        {
        }

        public LittleEndianInputStream(InputStream is1) : base(is1)
        {
        }

        /// <summary>
        /// <para>
        /// Reads up to <c>byte.length</c> bytes of data from this
        /// input stream into an array of bytes. This method blocks until some
        /// input is available.
        /// </para>
        /// <para>simulate java FilterInputStream</para>
        /// </summary>
        /// <param name="b"></param>
        /// <returns></returns>
        public override int Read(byte[] b)
        {
            return Read(b, 0, b.Length);
        }

        /// <summary>
        /// <para>
        /// Reads up to <c>len</c> bytes of data from this input stream
        /// into an array of bytes.If<c> len</c> is not zero, the method
        /// blocks until some input is available; otherwise, no
        /// bytes are read and<c>0</c> is returned.
        /// </para>
        /// <para>simulate java FilterInputStream</para>
        /// </summary>
        /// <param name="b"></param>
        /// <param name="off"></param>
        /// <param name="len"></param>
        /// <returns></returns>
        public override int Read(byte[] b, int off, int len)
        {
            return base.Read(b, off, len);
        }

        public override void Mark(int readlimit)
        {
            this.readLimit = readlimit;
            this.markPos = Position;
        }

        public override void Reset()
        {
            Seek(markPos - Position, SeekOrigin.Current);
        }

        public override long Skip(long n)
        {
            return Seek(n, SeekOrigin.Current);
        }

        public override int Available()
        {
            return (int)(Length - Position);
        }

        public override int ReadByte()
        {
            return (byte)ReadUByte();
        }

        public int ReadUByte()
        {
            int ch;
            try
            {
                ch = Read();
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
            CheckEOF(ch);
            return ch;
        }

        public double ReadDouble()
        {
            return BitConverter.Int64BitsToDouble(ReadLong());
        }

        public int ReadInt()
        {
            int ch1;
            int ch2;
            int ch3;
            int ch4;
            try
            {
                ch1 = Read();
                ch2 = Read();
                ch3 = Read();
                ch4 = Read();
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
            CheckEOF(ch1 | ch2 | ch3 | ch4);
            return (ch4 << 24) + (ch3 << 16) + (ch2 << 8) + (ch1 << 0);
        }

        public long ReadUInt()
        {
            long retNum = ReadInt();
            return retNum & 0x00FFFFFFFFL;
        }

        public long ReadLong()
        {
            int b0;
            int b1;
            int b2;
            int b3;
            int b4;
            int b5;
            int b6;
            int b7;
            try
            {
                b0 = Read();
                b1 = Read();
                b2 = Read();
                b3 = Read();
                b4 = Read();
                b5 = Read();
                b6 = Read();
                b7 = Read();
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
            CheckEOF(b0 | b1 | b2 | b3 | b4 | b5 | b6 | b7);
            return (((long)b7 << 56) +
                    ((long)b6 << 48) +
                    ((long)b5 << 40) +
                    ((long)b4 << 32) +
                    ((long)b3 << 24) +
                    (b2 << 16) +
                    (b1 << 8) +
                    (b0 << 0));
        }

        public short ReadShort()
        {
            return (short)ReadUShort();
        }

        public int ReadUShort()
        {
            int ch1;
            int ch2;
            try
            {
                ch1 = Read();
                ch2 = Read();
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
            CheckEOF(ch1 | ch2);
            return (ch2 << 8) + (ch1 << 0);
        }

        private static void CheckEOF(int value)
        {
            if (value < 0)
            {
                throw new RuntimeException("Unexpected end-of-file");
            }
        }

        public void ReadFully(byte[] buf)
        {
            ReadFully(buf, 0, buf.Length);
        }

        public void ReadFully(byte[] buf, int off, int len)
        {
            int max = off + len;
            for (int i = off; i < max; i++)
            {
                byte ch;
                try
                {
                    ch = (byte)Read();
                }
                catch (IOException e)
                {
                    throw new RuntimeException(e);
                }
                CheckEOF(ch);
                buf[i] = ch;
            }
        }
    }
}