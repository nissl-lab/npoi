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
    using NPOI.SS.Formula.Atp;
    using NPOI.SS.UserModel;

    [TestFixture]
    public class CountifsTests
    {

        [Test]
        public void TestCallFunction()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("test");
            IRow row1 = sheet.CreateRow(0);
            ICell cellA1 = row1.CreateCell(0, CellType.Formula);
            ICell cellB1 = row1.CreateCell(1, CellType.Numeric);
            ICell cellC1 = row1.CreateCell(2, CellType.Numeric);
            ICell cellD1 = row1.CreateCell(3, CellType.Numeric);
            ICell cellE1 = row1.CreateCell(4, CellType.Numeric);
            cellB1.SetCellValue(1);
            cellC1.SetCellValue(1);
            cellD1.SetCellValue(2);
            cellE1.SetCellValue(4);

            cellA1.SetCellFormula("COUNTIFS(B1:C1,1, D1:E1,2)");
            IFormulaEvaluator Evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            CellValue Evaluate = Evaluator.Evaluate(cellA1);
            Assert.AreEqual(1.0d, Evaluate.NumberValue);
        }

        // issue#825
        [Test]
        public void TestMultiRows()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("test");
            IRow row1 = sheet.CreateRow(0);
            IRow row2 = sheet.CreateRow(1);
            IRow row3 = sheet.CreateRow(2);

            ICell cellA1 = row1.CreateCell(0, CellType.Formula);

            ICell cellB1 = row2.CreateCell(1, CellType.Numeric);
            ICell cellC1 = row1.CreateCell(2, CellType.Numeric);
            ICell cellB2 = row2.CreateCell(1, CellType.Numeric);
            ICell cellC2 = row2.CreateCell(2, CellType.Numeric);
            ICell cellB3 = row3.CreateCell(1, CellType.Numeric);
            ICell cellC3 = row3.CreateCell(2, CellType.Numeric);


            cellB1.SetCellValue(1);
            cellB2.SetCellValue(2);
            cellB3.SetCellValue(2);

            cellC1.SetCellValue(2);
            cellC2.SetCellValue(2);
            cellC3.SetCellValue(3);

            // Only Row2 satisfy both conditions, so result should be 1
            cellA1.SetCellFormula("COUNTIFS(B1:B3,2, C1:C3,2)");
            IFormulaEvaluator Evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            CellValue Evaluate = Evaluator.Evaluate(cellA1);
            Assert.AreEqual(1.0d, Evaluate.NumberValue);
        }

        [Test]
        public void TestCallFunction_invalidArgs()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("test");
            IRow row1 = sheet.CreateRow(0);
            ICell cellA1 = row1.CreateCell(0, CellType.Formula);
            cellA1.CellFormula = (/*setter*/"COUNTIFS()");
            IFormulaEvaluator Evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            CellValue Evaluate = Evaluator.Evaluate(cellA1);
            Assert.AreEqual(15, Evaluate.ErrorValue);
            cellA1.CellFormula = (/*setter*/"COUNTIFS(A1:C1)");
            Evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            Evaluate = Evaluator.Evaluate(cellA1);
            Assert.AreEqual(15, Evaluate.ErrorValue);
            cellA1.SetCellFormula("COUNTIFS(A1:C1,2,2)");
            Evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            Evaluate = Evaluator.Evaluate(cellA1);
            Assert.AreEqual(15, Evaluate.ErrorValue);
        }
    }

}