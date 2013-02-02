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

    /// <summary>
    /// Adapts a plain byte array to <see cref="T:NPOI.Util.ILittleEndianInput"/>
    /// </summary>
    /// <remarks>@author Josh Micich</remarks>
    public class LittleEndianByteArrayInputStream : ILittleEndianInput
    {
        private byte[] _buf;
        private int _endIndex;
        private int _ReadIndex;

        public LittleEndianByteArrayInputStream(byte[] buf, int startOffset, int maxReadLen)
        {
            _buf = buf;
            _ReadIndex = startOffset;
            _endIndex = startOffset + maxReadLen;
        }
        public LittleEndianByteArrayInputStream(byte[] buf, int startOffset) :
            this(buf, startOffset, buf.Length - startOffset)
        {

        }
        public LittleEndianByteArrayInputStream(byte[] buf) :
            this(buf, 0, buf.Length)
        {

        }

        public int Available()
        {
            return _endIndex - _ReadIndex;
        }
        private void CheckPosition(int i)
        {
            if (i > _endIndex - _ReadIndex)
            {
                throw new RuntimeException("Buffer overrun");
            }
        }

        public int GetReadIndex()
        {
            return _ReadIndex;
        }
        public int ReadByte()
        {
            CheckPosition(1);
            return _buf[_ReadIndex++];
        }

        public int ReadInt()
        {
            CheckPosition(4);
            int i = _ReadIndex;

            int b0 = _buf[i++] & 0xFF;
            int b1 = _buf[i++] & 0xFF;
            int b2 = _buf[i++] & 0xFF;
            int b3 = _buf[i++] & 0xFF;
            _ReadIndex = i;
            return (b3 << 24) + (b2 << 16) + (b1 << 8) + (b0 << 0);
        }
        public long ReadLong()
        {
            CheckPosition(8);
            int i = _ReadIndex;

            int b0 = _buf[i++] & 0xFF;
            int b1 = _buf[i++] & 0xFF;
            int b2 = _buf[i++] & 0xFF;
            int b3 = _buf[i++] & 0xFF;
            int b4 = _buf[i++] & 0xFF;
            int b5 = _buf[i++] & 0xFF;
            int b6 = _buf[i++] & 0xFF;
            int b7 = _buf[i++] & 0xFF;
            _ReadIndex = i;
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
        public int ReadUByte()
        {
            CheckPosition(1);
            return _buf[_ReadIndex++] & 0xFF;
        }
        public int ReadUShort()
        {
            CheckPosition(2);
            int i = _ReadIndex;

            int b0 = _buf[i++] & 0xFF;
            int b1 = _buf[i++] & 0xFF;
            _ReadIndex = i;
            return (b1 << 8) + (b0 << 0);
        }
        public void ReadFully(byte[] buf, int off, int len)
        {
            CheckPosition(len);
            System.Array.Copy(_buf, _ReadIndex, buf, off, len);
            _ReadIndex += len;
        }
        public void ReadFully(byte[] buf)
        {
            ReadFully(buf, 0, buf.Length);
        }
        public double ReadDouble()
        {
            return BitConverter.Int64BitsToDouble(ReadLong());
        }
    }
}