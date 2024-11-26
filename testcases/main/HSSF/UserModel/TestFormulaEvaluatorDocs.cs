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

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.Collections;
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;
    using NPOI.SS.UserModel;

    /**
     * Tests to show that our documentation at
     *  http://poi.apache.org/hssf/eval.html
     * all actually works as we'd expect them to
     */
    [TestFixture]
    public class TestFormulaEvaluatorDocs
    {
        /**
         * http://poi.apache.org/hssf/eval.html#EvaluateAll
         */
        [Test]
        public void TestEvaluateAll()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet s1 = wb.CreateSheet();
            ISheet s2 = wb.CreateSheet();
            wb.SetSheetName(0, "S1");
            wb.SetSheetName(1, "S2");

            IRow s1r1 = s1.CreateRow(0);
            IRow s1r2 = s1.CreateRow(1);
            IRow s2r1 = s2.CreateRow(0);

            ICell s1r1c1 = s1r1.CreateCell(0);
            ICell s1r1c2 = s1r1.CreateCell(1);
            ICell s1r1c3 = s1r1.CreateCell(2);
            s1r1c1.SetCellValue(22.3);
            s1r1c2.SetCellValue(33.4);
            s1r1c3.CellFormula = ("SUM(A1:B1)");

            ICell s1r2c1 = s1r2.CreateCell(0);
            ICell s1r2c2 = s1r2.CreateCell(1);
            ICell s1r2c3 = s1r2.CreateCell(2);
            s1r2c1.SetCellValue(-1.2);
            s1r2c2.SetCellValue(-3.4);
            s1r2c3.CellFormula = ("SUM(A2:B2)");

            ICell s2r1c1 = s2r1.CreateCell(0);
            s2r1c1.CellFormula = ("S1!A1");

            // Not Evaluated yet
            Assert.AreEqual(0.0, s1r1c3.NumericCellValue, 0);
            Assert.AreEqual(0.0, s1r2c3.NumericCellValue, 0);
            Assert.AreEqual(0.0, s2r1c1.NumericCellValue, 0);

            // Do a full Evaluate, as per our docs
            // uses EvaluateFormulaCell()
            for (int sheetNum = 0; sheetNum < wb.NumberOfSheets; sheetNum++)
            {
                ISheet sheet = wb.GetSheetAt(sheetNum);
                HSSFFormulaEvaluator evaluator = new HSSFFormulaEvaluator(wb);

                for (IEnumerator rit = sheet.GetRowEnumerator(); rit.MoveNext(); )
                {
                    IRow r = (IRow)rit.Current;

                    for (IEnumerator cit = r.GetEnumerator(); cit.MoveNext(); )
                    {
                        ICell c = (ICell)cit.Current;
                        if (c.CellType == CellType.Formula)
                        {
                            evaluator.EvaluateFormulaCell(c);

                            // For Testing - all should be numeric
                            Assert.AreEqual(CellType.Numeric, evaluator.EvaluateFormulaCell(c));
                        }
                    }
                }
            }

            // Check now as expected
            Assert.AreEqual(55.7, wb.GetSheetAt(0).GetRow(0).GetCell(2).NumericCellValue, 0);
            Assert.AreEqual("SUM(A1:B1)", wb.GetSheetAt(0).GetRow(0).GetCell(2).CellFormula);
            Assert.AreEqual(CellType.Formula, wb.GetSheetAt(0).GetRow(0).GetCell(2).CellType);

            Assert.AreEqual(-4.6, wb.GetSheetAt(0).GetRow(1).GetCell(2).NumericCellValue, 0);
            Assert.AreEqual("SUM(A2:B2)", wb.GetSheetAt(0).GetRow(1).GetCell(2).CellFormula);
            Assert.AreEqual(CellType.Formula, wb.GetSheetAt(0).GetRow(1).GetCell(2).CellType);

            Assert.AreEqual(22.3, wb.GetSheetAt(1).GetRow(0).GetCell(0).NumericCellValue, 0);
            Assert.AreEqual("'S1'!A1", wb.GetSheetAt(1).GetRow(0).GetCell(0).CellFormula);
            Assert.AreEqual(CellType.Formula, wb.GetSheetAt(1).GetRow(0).GetCell(0).CellType);


            // Now do the alternate call, which zaps the formulas
            // uses EvaluateInCell()
            foreach (ISheet sheet in wb)
            {
                HSSFFormulaEvaluator evaluator = new HSSFFormulaEvaluator(wb);

                foreach (IRow r in sheet)
                {
                    foreach (ICell c in r)
                    {
                        if (c.CellType == CellType.Formula)
                        {
                            evaluator.EvaluateInCell(c);
                        }
                    }
                }
            }

            Assert.AreEqual(55.7, wb.GetSheetAt(0).GetRow(0).GetCell(2).NumericCellValue, 0);
            Assert.AreEqual(CellType.Numeric, wb.GetSheetAt(0).GetRow(0).GetCell(2).CellType);

            Assert.AreEqual(-4.6, wb.GetSheetAt(0).GetRow(1).GetCell(2).NumericCellValue, 0);
            Assert.AreEqual(CellType.Numeric, wb.GetSheetAt(0).GetRow(1).GetCell(2).CellType);

            Assert.AreEqual(22.3, wb.GetSheetAt(1).GetRow(0).GetCell(0).NumericCellValue, 0);
            Assert.AreEqual(CellType.Numeric, wb.GetSheetAt(1).GetRow(0).GetCell(0).CellType);
        }
    }
}