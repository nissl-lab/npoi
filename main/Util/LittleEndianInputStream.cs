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
    /// Wraps an <see cref="T:System.IO.Stream"/> providing <see cref="T:NPOI.Util.ILittleEndianInput"/><p/>
    /// 
    /// This class does not buffer any input, so the stream Read position maintained
    /// by this class is consistent with that of the inner stream.
    /// </summary>
    /// <remarks>
    /// @author Josh Micich
    /// </remarks>
    public class LittleEndianInputStream : ILittleEndianInput
    {
        Stream in1 = null;

        public int Available()
        {
            return (int)(in1.Length - in1.Position);
        }

        public LittleEndianInputStream(Stream is1)
        {
            in1 = is1;
        }
        public int ReadByte()
        {
            return (byte)ReadUByte();
        }
        public int ReadUByte()
        {
            int ch;
            try
            {
                ch = in1.ReadByte();
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
                ch1 = in1.ReadByte();
                ch2 = in1.ReadByte();
                ch3 = in1.ReadByte();
                ch4 = in1.ReadByte();
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
            CheckEOF(ch1 | ch2 | ch3 | ch4);
            return (ch4 << 24) + (ch3 << 16) + (ch2 << 8) + (ch1 << 0);
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
                b0 = in1.ReadByte();
                b1 = in1.ReadByte();
                b2 = in1.ReadByte();
                b3 = in1.ReadByte();
                b4 = in1.ReadByte();
                b5 = in1.ReadByte();
                b6 = in1.ReadByte();
                b7 = in1.ReadByte();
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
                ch1 = in1.ReadByte();
                ch2 = in1.ReadByte();
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
                    ch = (byte)in1.ReadByte();
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