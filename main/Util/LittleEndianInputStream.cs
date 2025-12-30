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
    /// <para>Wraps an <see cref="System.IO.Stream"/> providing <see cref="NPOI.Util.ILittleEndianInput"/><p/></para>
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
        private const int BUFFERED_SIZE = 8096;
        private int readLimit = -1;
        private long markPos = -1;

        public LittleEndianInputStream(Stream is1) : base(new FileInputStream(is1))
        {
        }

        public LittleEndianInputStream(InputStream is1) : base(is1.MarkSupported() ? is1 : new BufferedInputStream(is1, BUFFERED_SIZE))
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

        /// <summary>
        /// <para>
        /// Skips over and discards <c>n</c> bytes of data from the
        /// input stream. The <c>skip</c> method may, for a variety of
        /// reasons, end up skipping over some smaller number of bytes,
        /// possibly <c>0</c>. The actual number of bytes skipped is
        /// returned.
        /// </para>
        /// <para>
        /// This method simply performs <c>in.Skip(n)</c>.
        /// </para>
        /// </summary>
        /// <param name="n">  the number of bytes to be skipped.</param>
        /// <return>the actual number of bytes skipped.</return>
        /// <exception cref="IOException"> if <c>in.Skip(n)</c> throws an IOException.</exception>
        public override long Skip(long n)
        {
            var pos = input.Position;
            var newPos = Seek(n, SeekOrigin.Current);
            return newPos - pos;
        }

        public override int Available()
        {
            return base.Available();
        }

        public override int ReadByte()
        {
            return (byte)ReadUByte();
        }

        public int ReadUByte()
        {
            byte[] buf = new byte[1];
            try
            {
                CheckEOF(Read(buf), 1);
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
            
            return LittleEndian.GetUByte(buf);
        }

        public double ReadDouble()
        {
            return BitConverter.Int64BitsToDouble(ReadLong());
        }

        public int ReadInt()
        {
            byte[] buf = new byte[LittleEndianConsts.INT_SIZE];
            try
            {
                CheckEOF(Read(buf), buf.Length);
            }
            catch(IOException e)
            {
                throw new RuntimeException(e);
            }
            return LittleEndian.GetInt(buf);
        }

        public long ReadUInt()
        {
            long retNum = ReadInt();
            return retNum & 0x00FFFFFFFFL;
        }

        public long ReadLong()
        {
            byte[] buf = new byte[LittleEndianConsts.LONG_SIZE];
            try
            {
                CheckEOF(Read(buf), LittleEndianConsts.LONG_SIZE);
            }
            catch(IOException e)
            {
                throw new RuntimeException(e);
            }
            return LittleEndian.GetLong(buf);
        }

        public short ReadShort()
        {
            return (short)ReadUShort();
        }

        public int ReadUShort()
        {
            byte[] buf = new byte[LittleEndianConsts.SHORT_SIZE];
            try
            {
                CheckEOF(Read(buf), LittleEndianConsts.SHORT_SIZE);
            }
            catch(IOException e)
            {
                throw new RuntimeException(e);
            }
            return LittleEndian.GetUShort(buf);
        }

        private static void CheckEOF(int actualBytes, int expectedBytes)
        {
            if (expectedBytes != 0 && (actualBytes == -1 || actualBytes != expectedBytes))
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
            try
            {
                CheckEOF(Read(buf, off, len), len);
            }
            catch(IOException e)
            {
                throw new RuntimeException(e);
            }
        }

        public void ReadPlain(byte[] buf, int off, int len)
        {
            ReadFully(buf, off, len);
        }
    }
}