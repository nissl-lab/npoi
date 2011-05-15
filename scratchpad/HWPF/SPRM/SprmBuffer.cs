
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
        public SprmBuffer()
        {
            _buf = new byte[4];
            _offset = 0;
        }

        private int FindSprm(short opcode)
        {
            int operation = SprmOperation.GetOperationFromOpcode(opcode);
            int type = SprmOperation.GetTypeFromOpcode(opcode);

            SprmIterator si = new SprmIterator(_buf, 2);
            while (si.HasNext())
            {
                SprmOperation i = si.Next();
                if (i.Operation == operation && i.Type == type)
                    return i.GrpprlOffset;
            }
            return -1;
        }

        public void UpdateSprm(short opcode, byte operand)
        {
            int grpprlOffset = FindSprm(opcode);
            if (grpprlOffset != -1)
            {
                _buf[grpprlOffset] = operand;
                return;
            }
            else AddSprm(opcode, operand);
        }

        public void UpdateSprm(short opcode, short operand)
        {
            int grpprlOffset = FindSprm(opcode);
            if (grpprlOffset != -1)
            {
                LittleEndian.PutShort(_buf, grpprlOffset, operand);
                return;
            }
            else AddSprm(opcode, operand);
        }

        public void UpdateSprm(short opcode, int operand)
        {
            int grpprlOffset = FindSprm(opcode);
            if (grpprlOffset != -1)
            {
                LittleEndian.PutInt(_buf, grpprlOffset, operand);
                return;
            }
            else AddSprm(opcode, operand);
        }

        public void AddSprm(short opcode, byte operand)
        {
            int addition = LittleEndianConstants.SHORT_SIZE + LittleEndianConstants.BYTE_SIZE;
            EnsureCapacity(addition);
            LittleEndian.PutShort(_buf, _offset, opcode);
            _offset += LittleEndianConstants.SHORT_SIZE;
            _buf[_offset++] = operand;
        }
        public void AddSprm(short opcode, short operand)
        {
            int addition = LittleEndianConstants.SHORT_SIZE + LittleEndianConstants.SHORT_SIZE;
            EnsureCapacity(addition);
            LittleEndian.PutShort(_buf, _offset, opcode);
            _offset += LittleEndianConstants.SHORT_SIZE;
            LittleEndian.PutShort(_buf, _offset, operand);
            _offset += LittleEndianConstants.SHORT_SIZE;
        }
        public void AddSprm(short opcode, int operand)
        {
            int addition = LittleEndianConstants.SHORT_SIZE + LittleEndianConstants.INT_SIZE;
            EnsureCapacity(addition);
            LittleEndian.PutShort(_buf, _offset, opcode);
            _offset += LittleEndianConstants.SHORT_SIZE;
            LittleEndian.PutInt(_buf, _offset, operand);
            _offset += LittleEndianConstants.INT_SIZE;
        }
        public void AddSprm(short opcode, byte[] operand)
        {
            int addition = LittleEndianConstants.SHORT_SIZE + LittleEndianConstants.BYTE_SIZE + operand.Length;
            EnsureCapacity(addition);
            LittleEndian.PutShort(_buf, _offset, opcode);
            _offset += LittleEndianConstants.SHORT_SIZE;
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