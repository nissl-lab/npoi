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
     * @author aviks
     * @author Jason Height (jheight at chariot dot net dot au)
     * @author Danny Mui (dmui at apache dot org) (Leftover handling)
     */
    [Serializable]
    public class FuncPtg : AbstractFunctionPtg
    {

        public const byte sid = 0x21;
        public static int SIZE = 3;
        private int numParams = 0;

        /**Creates new function pointer from a byte array
         * usually called while Reading an excel file.
         */
        public FuncPtg(LittleEndianInput in1)
        {
            //field_1_num_args = data[ offset + 0 ];
            field_2_fnc_index = in1.ReadShort();

            FunctionMetadata fm = FunctionMetadataRegistry.GetFunctionByIndex(field_2_fnc_index);
            if (fm == null)
            {
                throw new Exception("Invalid built-in function index (" + field_2_fnc_index + ")");
            }
            numParams = fm.MinParams;
            returnClass = fm.ReturnClassCode;
            paramClass = fm.ParameterClassCodes;
        }
        public FuncPtg(int functionIndex)
        {
            field_2_fnc_index = (short)functionIndex;
            FunctionMetadata fm = FunctionMetadataRegistry.GetFunctionByIndex(functionIndex);
            numParams = fm.MinParams; // same as max since these are not var-arg funcs
            returnClass = fm.ReturnClassCode;
            paramClass = fm.ParameterClassCodes;
        }

        public override void Write(LittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteShort(field_2_fnc_index);
        }
        public override void WriteBytes(byte[] array, int offset)
        {
            array[offset + 0] = (byte)(sid + PtgClass);
            LittleEndian.PutShort(array, offset + 1, field_2_fnc_index);
        }

        public override int NumberOfOperands
        {
            get { return numParams; }
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
            sb.Append(" nArgs=").Append(numParams);
            sb.Append("]");
            return sb.ToString();
        }
    }
}