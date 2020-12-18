
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */


namespace NPOI.HWPF.SPRM
{
    using System;
    using System.Collections;
    using NPOI.Util;


    public class SprmBuffer : ICloneable
    {
        byte[] _buf;
        int _offset;
        bool _istd;
        private int _sprmsStartOffset;

        public SprmBuffer(byte[] buf, bool Istd)
        {
            _offset = buf.Length;
            _buf = buf;
            _istd = Istd;
        }
        public SprmBuffer(byte[] buf)
            : this(buf, false)
        {

        }
        public SprmBuffer(byte[] buf, bool istd, int sprmsStartOffset)
        {
            _offset = buf.Length;
            _buf = buf;
            _istd = istd;
            _sprmsStartOffset = sprmsStartOffset;
        }
        public SprmBuffer(byte[] buf, int _sprmsStartOffset)
            :this(buf, false, _sprmsStartOffset)
        {
            
        }

        public SprmBuffer(int sprmsStartOffset)
        {
            _buf = new byte[sprmsStartOffset + 4];
            _offset = sprmsStartOffset;
            _sprmsStartOffset = sprmsStartOffset;
        }

        public SprmBuffer()
        {
            _buf = new byte[4];
            _offset = 0;
        }

        internal SprmOperation FindSprm(short opcode)
        {
            int operation = SprmOperation.GetOperationFromOpcode(opcode);
            int type = SprmOperation.GetTypeFromOpcode(opcode);

            SprmIterator si = new SprmIterator(_buf, 2);
            while (si.HasNext())
            {
                SprmOperation i = si.Next();
                if (i.Operation == operation && i.Type == type)
                    return i;
            }
            return null;
        }


        public SprmIterator Iterator()
        {
            return new SprmIterator(_buf, _sprmsStartOffset);
        }

        internal int FindSprmOffset(short opcode)
        {
            SprmOperation sprmOperation = FindSprm(opcode);
            if (sprmOperation == null)
                return -1;

            return sprmOperation.GrpprlOffset;
        }
        public void UpdateSprm(short opcode, byte operand)
        {
            int grpprlOffset = FindSprmOffset(opcode);
            if (grpprlOffset != -1)
            {
                _buf[grpprlOffset] = operand;
                return;
            }
            else AddSprm(opcode, operand);
        }

        public void UpdateSprm(short opcode, short operand)
        {
            int grpprlOffset = FindSprmOffset(opcode);
            if (grpprlOffset != -1)
            {
                LittleEndian.PutShort(_buf, grpprlOffset, operand);
                return;
            }
            else AddSprm(opcode, operand);
        }

        public void UpdateSprm(short opcode, int operand)
        {
            int grpprlOffset = FindSprmOffset(opcode);
            if (grpprlOffset != -1)
            {
                LittleEndian.PutInt(_buf, grpprlOffset, operand);
                return;
            }
            else AddSprm(opcode, operand);
        }

        public void AddSprm(short opcode, byte operand)
        {
            int addition = LittleEndianConsts.SHORT_SIZE + LittleEndianConsts.BYTE_SIZE;
            EnsureCapacity(addition);
            LittleEndian.PutShort(_buf, _offset, opcode);
            _offset += LittleEndianConsts.SHORT_SIZE;
            _buf[_offset++] = operand;
        }
        public void AddSprm(short opcode, short operand)
        {
            int addition = LittleEndianConsts.SHORT_SIZE + LittleEndianConsts.SHORT_SIZE;
            EnsureCapacity(addition);
            LittleEndian.PutShort(_buf, _offset, opcode);
            _offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutShort(_buf, _offset, operand);
            _offset += LittleEndianConsts.SHORT_SIZE;
        }
        public void AddSprm(short opcode, int operand)
        {
            int addition = LittleEndianConsts.SHORT_SIZE + LittleEndianConsts.INT_SIZE;
            EnsureCapacity(addition);
            LittleEndian.PutShort(_buf, _offset, opcode);
            _offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutInt(_buf, _offset, operand);
            _offset += LittleEndianConsts.INT_SIZE;
        }
        public void AddSprm(short opcode, byte[] operand)
        {
            int addition = LittleEndianConsts.SHORT_SIZE + LittleEndianConsts.BYTE_SIZE + operand.Length;
            EnsureCapacity(addition);
            LittleEndian.PutShort(_buf, _offset, opcode);
            _offset += LittleEndianConsts.SHORT_SIZE;
            _buf[_offset++] = (byte)operand.Length;
            Array.Copy(operand, 0, _buf, _offset, operand.Length);
        }

        public byte[] ToByteArray()
        {
            return _buf;
        }

        public override bool Equals(Object obj)
        {
            SprmBuffer sprmBuf = (SprmBuffer)obj;
            return (Arrays.Equals(_buf, sprmBuf._buf));
        }

        public void Append(byte[] grpprl)
        {
            EnsureCapacity(grpprl.Length);
            System.Array.Copy(grpprl, 0, _buf, _offset, grpprl.Length);
        }
        public void Append(byte[] grpprl, int offset)
        {
            EnsureCapacity(grpprl.Length - offset);
            System.Array.Copy(grpprl, offset, _buf, _offset, grpprl.Length - offset);
            _offset += grpprl.Length - offset;
        }

        public Object Clone()
        {
            SprmBuffer retVal = new SprmBuffer();
            retVal._buf = new byte[_buf.Length];
            Array.Copy(_buf, 0, retVal._buf, 0, _buf.Length);
            return retVal;
        }

        private void EnsureCapacity(int Addition)
        {
            if (_offset + Addition >= _buf.Length)
            {
                // Add 6 more than they need for use the next iteration
                byte[] newBuf = new byte[_offset + Addition + 6];
                Array.Copy(_buf, 0, newBuf, 0, _buf.Length);
                _buf = newBuf;
            }
        }
    }
}