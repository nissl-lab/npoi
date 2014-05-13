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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;

namespace TestCases.SS.Formula.Atp
{
    [TestFixture]
    public class TestIfError
    {
        /**
     * =IFERROR(210/35,\"Error in calculation\")"  Divides 210 by 35 and returns 6.0
     * =IFERROR(55/0,\"Error in calculation\")"    Divides 55 by 0 and returns the error text
     * =IFERROR(C1,\"Error in calculation\")"      References the result of dividing 55 by 0 and returns the error text
     */
        [Test]
        public void TestEvaluate()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sh = wb.CreateSheet();
            IRow row1 = sh.CreateRow(0);
            IRow row2 = sh.CreateRow(1);

            // Create cells
            row1.CreateCell(0, CellType.Numeric);
            row1.CreateCell(1, CellType.Numeric);
            row1.CreateCell(2, CellType.Numeric);
            row2.CreateCell(0, CellType.Numeric);
            row2.CreateCell(1, CellType.Numeric);

            // Create references
            CellReference a1 = new CellReference("A1");
            CellReference a2 = new CellReference("A2");
            CellReference b1 = new CellReference("B1");
            CellReference b2 = new CellReference("B2");
            CellReference c1 = new CellReference("C1");

            // Set values
            sh.GetRow(a1.Row).GetCell(a1.Col).SetCellValue(210);
            sh.GetRow(a2.Row).GetCell(a2.Col).SetCellValue(55);
            sh.GetRow(b1.Row).GetCell(b1.Col).SetCellValue(35);
            sh.GetRow(b2.Row).GetCell(b2.Col).SetCellValue(0);
            sh.GetRow(c1.Row).GetCell(c1.Col).SetCellFormula("A1/B2");

            ICell cell1 = sh.CreateRow(3).CreateCell(0);
            cell1.SetCellFormula("IFERROR(A1/B1,\"Error in calculation\")");
            ICell cell2 = sh.CreateRow(3).CreateCell(0);
            cell2.SetCellFormula("IFERROR(A2/B2,\"Error in calculation\")");
            ICell cell3 = sh.CreateRow(3).CreateCell(0);
            cell3.SetCellFormula("IFERROR(C1,\"error\")");

            double accuracy = 1E-9;

            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            Assert.AreEqual(CellType.Numeric, evaluator.Evaluate(cell1).CellType, "Checks that the cell is numeric");
            Assert.AreEqual(6.0, evaluator.Evaluate(cell1).NumberValue, accuracy, "Divides 210 by 35 and returns 6.0");


            Assert.AreEqual(CellType.String, evaluator.Evaluate(cell2).CellType, "Checks that the cell is numeric");
            Assert.AreEqual("Error in calculation", evaluator.Evaluate(cell2).StringValue, "Rounds -10 to a nearest multiple of -3 (-9)");

            Assert.AreEqual(CellType.String, evaluator.Evaluate(cell3).CellType, "Check that C1 returns string");
            Assert.AreEqual("error", evaluator.Evaluate(cell3).StringValue, "Check that C1 returns string \"error\"");
        }
    }
}
