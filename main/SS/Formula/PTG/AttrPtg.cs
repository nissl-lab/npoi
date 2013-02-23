/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.SS.Formula.PTG
{
    using System;
    using System.Text;
    using NPOI.Util;
    


    /**
     * "Special Attributes"
     * This seems to be a Misc Stuff and Junk record.  One function it serves Is
     * in SUM functions (i.e. SUM(A1:A3) causes an area PTG then an ATTR with the SUM option Set)
     * @author  andy
     * @author Jason Height (jheight at chariot dot net dot au)
     */
    public class AttrPtg : ControlPtg
    {
        public const byte sid = 0x19;
        private const int SIZE = 4;
        private byte field_1_options;
        private short field_2_data;

        /** only used for tAttrChoose: table of offsets to starts of args */
        private int[] _jumpTable;
        /** only used for tAttrChoose: offset to the tFuncVar for CHOOSE() */
        private int _chooseFuncOffset;

        // flags 'volatile' and 'space', can be combined.  
        // OOO spec says other combinations are theoretically possible but not likely to occur.
        private static BitField semiVolatile = BitFieldFactory.GetInstance(0x01);
        private static BitField optiIf = BitFieldFactory.GetInstance(0x02);
        private static BitField optiChoose = BitFieldFactory.GetInstance(0x04);
        private static BitField optiSkip = BitFieldFactory.GetInstance(0x08); // skip
        private static BitField optiSum = BitFieldFactory.GetInstance(0x10);
        private static BitField baxcel = BitFieldFactory.GetInstance(0x20); // 'assignment-style formula in a macro sheet'
        private static BitField space = BitFieldFactory.GetInstance(0x40);

        public static readonly AttrPtg SUM = new AttrPtg(0x0010, 0, null, -1);

        public enum SpaceType
        {
            /** 00H = Spaces before the next token (not allowed before tParen token) */
            SpaceBefore = 0x00,
            /** 01H = Carriage returns before the next token (not allowed before tParen token) */
            CrBefore = 0x01,
            /** 02H = Spaces before opening parenthesis (only allowed before tParen token) */
            SpaceBeforeOpenParen = 0x02,
            /** 03H = Carriage returns before opening parenthesis (only allowed before tParen token) */
            CrBeforeOpenParen = 0x03,
            /** 04H = Spaces before closing parenthesis (only allowed before tParen, tFunc, and tFuncVar tokens) */
            SpaceBeforeCloseParen = 0x04,
            /** 05H = Carriage returns before closing parenthesis (only allowed before tParen, tFunc, and tFuncVar tokens) */
            CrBeforeCloseParen = 0x05,
            /** 06H = Spaces following the equality sign (only in macro sheets) */
            SpaceAfterEquality = 0x06
        }

        public AttrPtg()
        {
            _jumpTable = null;
            _chooseFuncOffset = -1;
        }

        public AttrPtg(ILittleEndianInput in1)
        {
            field_1_options =(byte)in1.ReadByte();
            field_2_data = in1.ReadShort();
            if (IsOptimizedChoose)
            {
                int nCases = field_2_data;
                int[] jumpTable = new int[nCases];
                for (int i = 0; i < jumpTable.Length; i++)
                {
                    jumpTable[i] = in1.ReadUShort();
                }
                _jumpTable = jumpTable;
                _chooseFuncOffset = in1.ReadUShort();
            }
            else
            {
                _jumpTable = null;
                _chooseFuncOffset = -1;
            }

        }
        private AttrPtg(int options, int data, int[] jt, int chooseFuncOffset)
        {
            field_1_options = (byte)options;
            field_2_data = (short)data;
            _jumpTable = jt;
            _chooseFuncOffset = chooseFuncOffset;
        }

        /// <summary>
        /// Creates the space.
        /// </summary>
        /// <param name="type">a constant from SpaceType</param>
        /// <param name="count">The count.</param>
        public static AttrPtg CreateSpace(SpaceType type, int count)
        {
            int data = ((int) type) & 0x00FF | (count << 8) & 0x00FFFF;
            return new AttrPtg(space.Set(0), data, null, -1);
        }

        /// <summary>
        /// Creates if.
        /// </summary>
        /// <param name="dist">distance (in bytes) to start of either
        /// tFuncVar(IF) token (when false parameter is not present).</param>
        public static AttrPtg CreateIf(int dist)
        {
            return new AttrPtg(optiIf.Set(0), dist, null, -1);
        }

        /// <summary>
        /// Creates the skip.
        /// </summary>
        /// <param name="dist">distance (in bytes) to position behind tFuncVar(IF) token (minus 1).</param>
        public static AttrPtg CreateSkip(int dist)
        {
            return new AttrPtg(optiSkip.Set(0), dist, null, -1);
        }
        public static AttrPtg GetSumSingle()
        {
            return new AttrPtg(optiSum.Set(0), 0, null, -1);
        }

        public bool IsSemiVolatile
        {
            get { return semiVolatile.IsSet(field_1_options); }
        }

        public bool IsOptimizedIf
        {
            get { return optiIf.IsSet(field_1_options); }
            set { field_1_options = optiIf.SetByteBoolean(field_1_options, value); }
        }

        public bool IsOptimizedChoose
        {
            get { return optiChoose.IsSet(field_1_options); }
        }

        public bool IsSum
        {
            get { return optiSum.IsSet(field_1_options); }
            set { field_1_options = optiSum.SetByteBoolean(field_1_options, value); }
        }

        // lets hope no one uses this anymore
        public bool IsBaxcel
        {
            get{return baxcel.IsSet(field_1_options);}
        }

        // biff3&4 only  shouldn't happen anymore
        public bool IsSpace
        {
            get { return space.IsSet(field_1_options); }
        }
        public bool IsSkip
        {
            get { return optiSkip.IsSet(field_1_options); }
        }

        public short Data
        {
            get { return field_2_data; }
            set { field_2_data = value; }
        }
        public int[] JumpTable
        {
            get
            {
                return (int[])_jumpTable.Clone();
            }
        }
        public int ChooseFuncOffset
        {
            get
            {
                if (_jumpTable == null)
                {
                    throw new InvalidOperationException("Not tAttrChoose");
                }
                return _chooseFuncOffset;
            }
        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");

            if (IsSemiVolatile)
            {
                sb.Append("volatile ");
            }
            if (IsSpace)
            {
                sb.Append("space count=").Append((field_2_data >> 8) & 0x00FF);
                sb.Append(" type=").Append(field_2_data & 0x00FF).Append(" ");
            }
            // the rest seem to be mutually exclusive
            if (IsOptimizedIf)
            {
                sb.Append("if dist=").Append(Data);
            }
            else if (IsOptimizedChoose)
            {
                sb.Append("choose nCases=").Append(Data);
            }
            else if (IsSkip)
            {
                sb.Append("skip dist=").Append(Data);
            }
            else if (IsSum)
            {
                sb.Append("sum ");
            }
            else if (IsBaxcel)
            {
                sb.Append("assign ");
            }
            sb.Append("]");
            return sb.ToString();
        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteByte(field_1_options);
            out1.WriteShort(field_2_data);
            int[] jt = _jumpTable;
            if (jt != null)
            {
                for (int i = 0; i < jt.Length; i++)
                {
                    out1.WriteShort(jt[i]);
                }
                out1.WriteShort(_chooseFuncOffset);
            }


        }


        public override int Size
        {
            get
            {
                if (_jumpTable != null)
                {
                    return SIZE + (_jumpTable.Length + 1) * LittleEndianConsts.SHORT_SIZE;
                }
               return SIZE;
            }
        }

        public String ToFormulaString(String[] operands)
        {
            if (space.IsSet(field_1_options))
            {
                return operands[0];
            }
            else if (optiIf.IsSet(field_1_options))
            {
                return ToFormulaString() + "(" + operands[0] + ")";
            }
            else if (optiSkip.IsSet(field_1_options))
            {
                return ToFormulaString() + operands[0];   //goto Isn't a real formula element should not show up
            }
            else
            {
                return ToFormulaString() + "(" + operands[0] + ")";
            }
        }


        public int NumberOfOperands
        {
            get { return 1; }
        }

        public int Type
        {
            get { return -1; }
        }

        public override String ToFormulaString()
        {
            if (semiVolatile.IsSet(field_1_options))
            {
                return "ATTR(semiVolatile)";
            }
            if (optiIf.IsSet(field_1_options))
            {
                return "IF";
            }
            if (optiChoose.IsSet(field_1_options))
            {
                return "CHOOSE";
            }
            if (optiSkip.IsSet(field_1_options))
            {
                return "";
            }
            if (optiSum.IsSet(field_1_options))
            {
                return "SUM";
            }
            if (baxcel.IsSet(field_1_options))
            {
                return "ATTR(baxcel)";
            }
            if (space.IsSet(field_1_options))
            {
                return "";
            }
            return "UNKNOWN ATTRIBUTE";
        }

        public override Object Clone()
        {
            int[] jt;
            if (_jumpTable == null)
            {
                jt = null;
            }
            else
            {
                jt = (int[])_jumpTable.Clone();
            }
            return new AttrPtg(field_1_options, field_2_data, jt, _chooseFuncOffset);
        }
    }
}
