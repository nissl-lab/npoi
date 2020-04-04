/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

namespace TestCases.XSSF.Streaming
{
    using NPOI.SS.UserModel;
    using NPOI.XSSF;
    using NPOI.XSSF.Streaming;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;

    /**
     * Formula Evaluation with SXSSF.
     * 
     * Note that SXSSF can only Evaluate formulas where the
     *  cell is in the current window, and all references
     *  from the cell are in the current window
     */
    [TestFixture]
    public class TestSXSSFFormulaEvaluation
    {
        public static SXSSFITestDataProvider _testDataProvider = SXSSFITestDataProvider.instance;

        /**
         * EvaluateAll will normally fail, as any reference or
         *  formula outside of the window will fail, and any
         *  non-active sheets will fail
         */
        [Test]

        public void TestEvaluateAllFails()
        {
            SXSSFWorkbook wb = new SXSSFWorkbook(5);
            SXSSFSheet s = wb.CreateSheet() as SXSSFSheet;

            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();

            s.CreateRow(0).CreateCell(0).CellFormula = (/*setter*/"1+2");
            s.CreateRow(1).CreateCell(0).CellFormula = (/*setter*/"A21");
            for (int i = 2; i < 19; i++) { s.CreateRow(i); }

            // Cells outside window will fail, whether referenced or not
            s.CreateRow(19).CreateCell(0).CellFormula = (/*setter*/"A1+A2");
            s.CreateRow(20).CreateCell(0).CellFormula = (/*setter*/"A1+A11+100");
            try
            {
                eval.EvaluateAll();
                Assert.Fail("Evaluate All shouldn't work, as some cells outside the window");
            }
            catch (RowFlushedException)
            {
                // Expected
            }


            // Inactive sheets will fail
            XSSFWorkbook xwb = new XSSFWorkbook();
            xwb.CreateSheet("Open");
            xwb.CreateSheet("Closed");

            wb.Close();
            wb = new SXSSFWorkbook(xwb, 5);
            s = wb.GetSheet("Closed") as SXSSFSheet;
            s.FlushRows();
            s = wb.GetSheet("Open") as SXSSFSheet;
            s.CreateRow(0).CreateCell(0).CellFormula = (/*setter*/"1+2");

            eval = wb.GetCreationHelper().CreateFormulaEvaluator();
            try
            {
                eval.EvaluateAll();
                Assert.Fail("Evaluate All shouldn't work, as sheets flushed");
            }
            catch (SheetsFlushedException) { }

            wb.Close();
        }

        [Test]
        public void TestEvaluateRefOutsideWindowFails()
        {
            SXSSFWorkbook wb = new SXSSFWorkbook(5);
            SXSSFSheet s = wb.CreateSheet() as SXSSFSheet;

            s.CreateRow(0).CreateCell(0).CellFormula = (/*setter*/"1+2");
            Assert.AreEqual(false, s.AllRowsFlushed);
            Assert.AreEqual(-1, s.LastFlushedRowNumber);

            for (int i = 1; i <= 19; i++) { s.CreateRow(i); }
            ICell c = s.CreateRow(20).CreateCell(0);
            c.CellFormula = (/*setter*/"A1+100");

            Assert.AreEqual(false, s.AllRowsFlushed);
            Assert.AreEqual(15, s.LastFlushedRowNumber);

            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();
            try
            {
                eval.EvaluateFormulaCell(c);
                Assert.Fail("Evaluate shouldn't work, as reference outside the window");
            }
            catch (RowFlushedException)
            {
                // Expected
            }

            wb.Close();
        }

        /**
         * If all formula cells + their references are inside the window,
         *  then Evaluation works
         */
        [Test]
        public void TestEvaluateAllInWindow()
        {
            SXSSFWorkbook wb = new SXSSFWorkbook(5);
            SXSSFSheet s = wb.CreateSheet() as SXSSFSheet;
            s.CreateRow(0).CreateCell(0).CellFormula = (/*setter*/"1+2");
            s.CreateRow(1).CreateCell(1).CellFormula = (/*setter*/"A1+10");
            s.CreateRow(2).CreateCell(2).CellFormula = (/*setter*/"B2+100");

            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();
            eval.EvaluateAll();

            Assert.AreEqual(3, (int)s.GetRow(0).GetCell(0).NumericCellValue);
            Assert.AreEqual(13, (int)s.GetRow(1).GetCell(1).NumericCellValue);
            Assert.AreEqual(113, (int)s.GetRow(2).GetCell(2).NumericCellValue);

            wb.Close();
        }

        [Test]
        public void TestEvaluateRefInsideWindow()
        {
            SXSSFWorkbook wb = new SXSSFWorkbook(5);
            SXSSFSheet s = wb.CreateSheet() as SXSSFSheet;

            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();

            SXSSFCell c = s.CreateRow(0).CreateCell(0) as SXSSFCell;
            c.SetCellValue(1.5);

            c = s.CreateRow(1).CreateCell(0) as SXSSFCell;
            c.CellFormula = (/*setter*/"A1*2");

            Assert.AreEqual(0, (int)c.NumericCellValue);
            eval.EvaluateFormulaCell(c);
            Assert.AreEqual(3, (int)c.NumericCellValue);

            wb.Close();
        }

        [Test]
        public void TestEvaluateSimple()
        {
            SXSSFWorkbook wb = new SXSSFWorkbook(5);
            SXSSFSheet s = wb.CreateSheet() as SXSSFSheet;

            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();

            SXSSFCell c = s.CreateRow(0).CreateCell(0) as SXSSFCell;
            c.CellFormula = (/*setter*/"1+2");
            Assert.AreEqual(0, (int)c.NumericCellValue);
            eval.EvaluateFormulaCell(c);
            Assert.AreEqual(3, (int)c.NumericCellValue);

            c = s.CreateRow(1).CreateCell(0) as SXSSFCell;
            c.CellFormula = (/*setter*/"CONCATENATE(\"hello\",\" \",\"world\")");
            eval.EvaluateFormulaCell(c);
            Assert.AreEqual("hello world", c.StringCellValue);


        }
    }

}