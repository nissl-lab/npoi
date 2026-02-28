/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.Collections;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.SS.Formula;
    using NPOI.SS.UserModel;
    using NPOI.SS.Formula.PTG;

    /**
     * Test utility class to Get <c>Ptg</c> arrays out of formula cells
     * 
     * @author Josh Micich
     */
    public class FormulaExtractor
    {

        private FormulaExtractor()
        {
            // no instances of this class
        }
        
        public static Ptg[] GetPtgs(ICell cell)
        {
            CellValueRecordInterface vr = ((HSSFCell)cell).CellValueRecord;
            if (!(vr is FormulaRecordAggregate))
            {
                throw new ArgumentException("Not a formula cell");
            }
            FormulaRecordAggregate fra = (FormulaRecordAggregate)vr;
            return fra.FormulaRecord.ParsedExpression;
        }

    }
}