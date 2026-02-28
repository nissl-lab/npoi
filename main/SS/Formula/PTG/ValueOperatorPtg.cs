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

    //import org.apache.poi.hssf.usermodel.HSSFWorkbook;

    /**
     * Common baseclass of all value operators.
     * Subclasses include all Unary and binary operators except for the reference operators (IntersectionPtg, RangePtg, UnionPtg) 
     * 
     * @author Josh Micich
     */
    public abstract class ValueOperatorPtg : OperationPtg
    {

        /**
         * All Operator <c>Ptg</c>s are base tokens (i.e. are not RVA classified)  
         */
        public override bool IsBaseToken
        {
            get { return true; }
        }

        public override byte DefaultOperandClass
        {
            get { return Ptg.CLASS_VALUE; }
        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(Sid + PtgClass);
        }

        protected abstract byte Sid { get; }

        public override int Size
        {
            get { return 1; }
        }

        public override String ToFormulaString() 
        {
    	    throw new NotImplementedException("ToFormulaString(String[] operands) should be used for subclasses of OperationPtgs");
	    }
    }
}