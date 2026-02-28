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
    using NPOI.Util;
    using NPOI.SS.Formula.Function;


    /**
     *
     * @author Jason Height (jheight at chariot dot net dot au)
     */
    public class FuncVarPtg : AbstractFunctionPtg
    {
        public const byte sid = 0x22;
        private const int SIZE = 4;

        /**
 * Single instance of this token for 'sum() taking a single argument'
 */
        public static readonly OperationPtg SUM = FuncVarPtg.Create("SUM", 1);

        private FuncVarPtg(int functionIndex, int returnClass, byte[] paramClasses, int numArgs)
            : base(functionIndex, returnClass, paramClasses, numArgs)
        {

        }

        /**Creates new function pointer from a byte array
 * usually called while reading an excel file.
 */
        public static FuncVarPtg Create(ILittleEndianInput in1)
        {
            return Create(in1.ReadByte(), in1.ReadShort());
        }

        /**
         * Create a function ptg from a string tokenised by the parser
         */
        public static FuncVarPtg Create(String pName, int numArgs)
        {
            return Create(numArgs, LookupIndex(pName));
        }

        private static FuncVarPtg Create(int numArgs, int functionIndex)
        {
            FunctionMetadata fm = FunctionMetadataRegistry.GetFunctionByIndex(functionIndex);
            if (fm == null)
            {
                // Happens only as a result of a call to FormulaParser.parse(), with a non-built-in function name
                return new FuncVarPtg(functionIndex, Ptg.CLASS_VALUE, new byte[] { Ptg.CLASS_VALUE }, numArgs);
            }
            return new FuncVarPtg(functionIndex, fm.ReturnClassCode, fm.ParameterClassCodes, numArgs);
        }


        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteByte(_numberOfArgs);
            out1.WriteShort(_functionIndex);
        }

        public override int Size
        {
            get { return SIZE; }
        }
    }
}