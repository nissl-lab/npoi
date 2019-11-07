/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.SS.Formula.Functions
{

    using System;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    /**
     * Tests for {@link Financed}
     * 
     * @author Torstein Svendsen (torstei@officenet.no)
     */
    [TestFixture]
    public class TestFind
    {
        [Test]
        public void TestFind1()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

            ConfirmResult(fe, cell, "Find(\"h\", \"haystack\")", 1);
            ConfirmResult(fe, cell, "Find(\"a\", \"haystack\",2)", 2);
            ConfirmResult(fe, cell, "Find(\"a\", \"haystack\",3)", 6);

            // number args Converted to text
            ConfirmResult(fe, cell, "Find(7, 32768)", 3);
            ConfirmResult(fe, cell, "Find(\"34\", 1341235233412, 3)", 10);
            ConfirmResult(fe, cell, "Find(5, 87654)", 4);

            // Errors
            ConfirmError(fe, cell, "Find(\"n\", \"haystack\")", FormulaError.VALUE);
            ConfirmError(fe, cell, "Find(\"k\", \"haystack\",9)", FormulaError.VALUE);
            ConfirmError(fe, cell, "Find(\"k\", \"haystack\",#REF!)", FormulaError.REF);
            ConfirmError(fe, cell, "Find(\"k\", \"haystack\",0)", FormulaError.VALUE);
            ConfirmError(fe, cell, "Find(#DIV/0!, #N/A, #REF!)", FormulaError.DIV0);
            ConfirmError(fe, cell, "Find(2, #N/A, #REF!)", FormulaError.NA);
        }

        private static void ConfirmResult(HSSFFormulaEvaluator fe, ICell cell, String formulaText,
                int expectedResult)
        {
            cell.CellFormula=(formulaText);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(result.CellType, CellType.Numeric);
            Assert.AreEqual(expectedResult, result.NumberValue, 0.0);
        }

        private static void ConfirmError(HSSFFormulaEvaluator fe, ICell cell, String formulaText,
                FormulaError expectedErrorCode)
        {
            cell.CellFormula=(formulaText);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(result.CellType, CellType.Error);
            Assert.AreEqual(expectedErrorCode.Code, result.ErrorValue);
        }
    }

}