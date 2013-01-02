
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

    /**
     * This class Is used to represent a sprm operation from a Word 97/2000/XP
     * document.
     * @author Ryan Ackley
     * @version 1.0
     */
    public class SprmOperation
    {
        static private BitField OP_BITFIELD = BitFieldFactory.GetInstance(0x1ff);
        static private BitField SPECIAL_BITFIELD = BitFieldFactory.GetInstance(0x200);
        static private BitField TYPE_BITFIELD = BitFieldFactory.GetInstance(0x1c00);
        static private BitField SIZECODE_BITFIELD = BitFieldFactory.GetInstance(0xe000);

        static private short LONG_SPRM_TABLE = unchecked((short)0xd608);
        static private short LONG_SPRM_PARAGRAPH = unchecked((short)0xc615);

        public const int TYPE_PAP = 1;
        public const int TYPE_CHP = 2;
        public const int TYPE_PIC = 3;
        public const int TYPE_SEP = 4;
        public const int TYPE_TAP = 5;

        private int _type;
        private int _operation;
        private int _gOffset;
        private byte[] _grpprl;
        private int _sizeCode;
        private int _size;

        public SprmOperation(byte[] grpprl, int offset)
        {
            _grpprl = grpprl;

            short sprmStart = LittleEndian.GetShort(grpprl, offset);

            _gOffset = offset + 2;

            _operation = OP_BITFIELD.GetValue(sprmStart);
            _type = TYPE_BITFIELD.GetValue(sprmStart);
            _sizeCode = SIZECODE_BITFIELD.GetValue(sprmStart);
            _size = InitSize(sprmStart);
        }

        public static int GetOperationFromOpcode(short opcode)
        {
            return OP_BITFIELD.GetValue(opcode);
        }

        public static int GetTypeFromOpcode(short opcode)
        {
            return TYPE_BITFIELD.GetValue(opcode);
        }
        byte[] _buf;
        public SprmOperation FindSprm(short opcode)
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
        public int Type
        {
            get
            {
                return _type;
            }
        }

        public int Operation
        {
            get
            {
                return _operation;
            }
        }

        public int GrpprlOffset
        {
            get
            {
                return _gOffset;
            }
        }
        public int Operand
        {
            get
            {
                switch (_sizeCode)
                {
                    case 0:
                    case 1:
                        return _grpprl[_gOffset];
                    case 2:
                    case 4:
                    case 5:
                        return LittleEndian.GetShort(_grpprl, _gOffset);
                    case 3:
                        return LittleEndian.GetInt(_grpprl, _gOffset);
                    case 6:
                        byte operandLength = _grpprl[_gOffset + 1];   //surely shorter than an int...

                        byte[] codeBytes = new byte[LittleEndianConsts.INT_SIZE]; //initialized to zeros by JVM
                        for (int i = 0; i < operandLength; i++)
                            if (_gOffset + i < _grpprl.Length)
                                codeBytes[i] = _grpprl[_gOffset + 1 + i];

                        return LittleEndian.GetInt(codeBytes, 0);
                    case 7:
                        byte[] threeByteInt = new byte[4];
                        threeByteInt[0] = _grpprl[_gOffset];
                        threeByteInt[1] = _grpprl[_gOffset + 1];
                        threeByteInt[2] = _grpprl[_gOffset + 2];
                        threeByteInt[3] = (byte)0;
                        return LittleEndian.GetInt(threeByteInt, 0);
                    default:
                        throw new ArgumentException("SPRM contains an invalid size code");
                }
            }
        }
        public int SizeCode
        {
            get
            {
                return _sizeCode;
            }
        }

        public int Size
        {
            get
            {
                return _size;
            }
        }

        public byte[] Grpprl
        {
            get
            {
                return _grpprl;
            }
        }
        private int InitSize(short sprm)
        {
            switch (_sizeCode)
            {
                case 0:
                case 1:
                    return 3;
                case 2:
                case 4:
                case 5:
                    return 4;
                case 3:
                    return 6;
                case 6:
                    if (sprm == LONG_SPRM_TABLE || sprm == LONG_SPRM_PARAGRAPH)
                    {
                        int retVal = (0x0000ffff & LittleEndian.GetShort(_grpprl, _gOffset)) + 3;
                        _gOffset += 2;
                        return retVal;
                    }
                    else
                    {
                        return (0x000000ff & _grpprl[_gOffset++]) + 3;
                    }
                case 7:
                    return 5;
                default:
                    throw new ArgumentException("SPRM contains an invalid size code");
            }
        }
    }
}
