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

namespace NPOI.HSSF.Record.Formula
{
    using System;
    using System.Text;
    using NPOI.Util;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Formula.Function;
    using NPOI.Util.IO;

    /**
     *
     * @author Jason Height (jheight at chariot dot net dot au)
     */
    public class FuncVarPtg : AbstractFunctionPtg
    {

        public const byte sid = 0x22;
        private static int SIZE = 4;

        /**Creates new function pointer from a byte array
         * usually called while Reading an excel file.
         */
        public FuncVarPtg(LittleEndianInput in1)
        {
            field_1_num_args = (byte)in1.ReadByte();
            field_2_fnc_index = in1.ReadShort();
            FunctionMetadata fm = FunctionMetadataRegistry.GetFunctionByIndex(field_2_fnc_index);
            if (fm == null)
            {
                // Happens only as a result of a call to FormulaParser.Parse(), with a non-built-in function name
                returnClass = Ptg.CLASS_VALUE;
                paramClass = new byte[] { Ptg.CLASS_VALUE };
            }
            else
            {
                returnClass = fm.ReturnClassCode;
                paramClass = fm.ParameterClassCodes;
            }
        }

        /**
         * Create a function ptg from a string tokenised by the Parser
         */
        public FuncVarPtg(String pName, byte pNumOperands)
        {
            field_1_num_args = pNumOperands;
            field_2_fnc_index = LookupIndex(pName);
            FunctionMetadata fm = FunctionMetadataRegistry.GetFunctionByIndex(field_2_fnc_index);
            if (fm == null)
            {
                // Happens only as a result of a call to FormulaParser.Parse(), with a non-built-in function name
                returnClass = Ptg.CLASS_VALUE;
                paramClass = new byte[] { Ptg.CLASS_VALUE };
            }
            else
            {
                returnClass = fm.ReturnClassCode;
                paramClass = fm.ParameterClassCodes;
            }
        }

        public override void Write(LittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteByte(field_1_num_args);
            out1.WriteShort(field_2_fnc_index);
        }

        public override void WriteBytes(byte[] array, int offset)
        {
            array[offset + 0] = (byte)(sid + PtgClass);
            array[offset + 1] = field_1_num_args;
            LittleEndian.PutShort(array, offset + 2, field_2_fnc_index);
        }

        public override int NumberOfOperands
        {
            get { return field_1_num_args; }
        }

        public override int Size
        {
            get { return SIZE; }
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append(LookupName(field_2_fnc_index));
            sb.Append(" nArgs=").Append(field_1_num_args);
            sb.Append("]");
            return sb.ToString();
        }
    }
}