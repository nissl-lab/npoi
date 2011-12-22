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

namespace TestCases.SS.Formula.Eval
{
    using System;

    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.SS.Formula.Eval;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    /**
     * Tests HSSFFormulaEvaluator for its handling of cell formula circular references.
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestCircularReferences
    {
        /**
         * Translates StackOverflowError into AssertFailedException
         */
        private static NPOI.SS.UserModel.CellValue EvaluateWithCycles(HSSFWorkbook wb,  ICell testCell)
        {
            HSSFFormulaEvaluator evaluator = new HSSFFormulaEvaluator(wb);
            try
            {
                return evaluator.Evaluate(testCell);
            }
            catch (StackOverflowException)
            {
                throw new AssertFailedException("circular reference caused stack overflow error");
            }
        }
        /**
         * Makes sure that the specified evaluated cell value represents a circular reference error.
         */
        private static void ConfirmCycleErrorCode(NPOI.SS.UserModel.CellValue cellValue)
        {
            Assert.IsTrue(cellValue.CellType == NPOI.SS.UserModel.CellType.ERROR);
            Assert.AreEqual((byte)ErrorEval.CIRCULAR_REF_ERROR.ErrorCode, cellValue.ErrorValue);
        }


        /**
         * ASF Bugzilla Bug 44413  
         * "INDEX() formula cannot contain its own location in the data array range" 
         */
        [TestMethod]
        public void TestIndexFormula()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Sheet1");

            short colB = 1;
            sheet.CreateRow(0).CreateCell(colB).SetCellValue(1);
            sheet.CreateRow(1).CreateCell(colB).SetCellValue(2);
            sheet.CreateRow(2).CreateCell(colB).SetCellValue(3);
            IRow row4 = sheet.CreateRow(3);
            ICell testCell = row4.CreateCell((short)0);
            // This formula should evaluate to the contents of B2,
            testCell.CellFormula = ("INDEX(A1:B4,2,2)");
            // However the range A1:B4 also includes the current cell A4.  If the other parameters
            // were 4 and 1, this would represent a circular reference.  Since POI 'fully' evaluates
            // arguments before invoking operators, POI must handle such potential cycles gracefully.


            NPOI.SS.UserModel.CellValue cellValue = EvaluateWithCycles(wb, testCell);

            Assert.IsTrue(cellValue.CellType == NPOI.SS.UserModel.CellType.NUMERIC);
            Assert.AreEqual(2, cellValue.NumberValue, 0);
        }

        /**
         * Cell A1 has formula "=A1"
         */
        [TestMethod]
        public void TestSimpleCircularReference()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Sheet1");

            IRow row = sheet.CreateRow(0);
            ICell testCell = row.CreateCell(0);
            testCell.CellFormula = ("A1");

            NPOI.SS.UserModel.CellValue cellValue = EvaluateWithCycles(wb,testCell);

            ConfirmCycleErrorCode(cellValue);
        }

        /**
         * A1=B1, B1=C1, C1=D1, D1=A1
         */
        [TestMethod]
        public void TestMultiLevelCircularReference()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Sheet1");

            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).CellFormula = ("B1");
            row.CreateCell(1).CellFormula = ("C1");
            row.CreateCell(2).CellFormula = ("D1");
            ICell testCell = row.CreateCell(3);
            testCell.CellFormula = ("A1");

            NPOI.SS.UserModel.CellValue cellValue = EvaluateWithCycles(wb,  testCell);

            ConfirmCycleErrorCode(cellValue);
        }
    }
}