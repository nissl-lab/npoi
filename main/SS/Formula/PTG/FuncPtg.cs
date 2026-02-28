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
    using NPOI.HSSF.Record;
    
    using NPOI.SS.Formula.Function;


    /**
     * @author aviks
     * @author Jason Height (jheight at chariot dot net dot au)
     * @author Danny Mui (dmui at apache dot org) (Leftover handling)
     */
    [Serializable]
    public class FuncPtg : AbstractFunctionPtg
    {

        public const byte sid = 0x21;
        public const int SIZE = 3;
        // not used: private int numParams = 0;
    public static FuncPtg Create(ILittleEndianInput in1) {
        return Create(in1.ReadUShort());
    }
    private FuncPtg(int funcIndex, FunctionMetadata fm):
        base(funcIndex, fm.ReturnClassCode, fm.ParameterClassCodes, fm.MinParams)  // minParams same as max since these are not var-arg funcs {
    {
    }
    public static FuncPtg Create(int functionIndex) {
        FunctionMetadata fm = FunctionMetadataRegistry.GetFunctionByIndex(functionIndex);
        if(fm == null) {
            throw new Exception("Invalid built-in function index (" + functionIndex + ")");
        }
        return new FuncPtg(functionIndex, fm);
    }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteShort(_functionIndex);
        }
        public override int Size
        {
            get { return SIZE; }
        }
    }
}