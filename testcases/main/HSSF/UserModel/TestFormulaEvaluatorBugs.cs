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
    using System.Configuration;
    using System.IO;
    using NUnit.Framework;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;
    using TestCases.HSSF;
    using TestCases.SS.Formula;
    /**
     * 
     */
    [TestFixture]
    public class TestFormulaEvaluatorBugs
    {
        private static bool OUTPUT_TEST_FILES = false;
        private String tmpDirName;

        [SetUp]
        public void SetUp()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            tmpDirName = ConfigurationManager.AppSettings["java.io.tmpdir"];
        }

        /**
         * An odd problem with EvaluateFormulaCell giving the
         *  right values when file is Opened, but changes
         *  to the source data in some versions of excel 
         *  doesn't cause them to be updated. However, other
         *  versions of excel, and gnumeric, work just fine
         * WARNING - tedious bug where you actually have to
         *  Open up excel
         */
        [Test]
        public void Test44636()
        {
            // Open the existing file, tweak one value and
            // re-calculate

            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("44636.xls");
            NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(0);
            IRow row = sheet.GetRow(0);

            row.GetCell(0).SetCellValue(4.2);
            row.GetCell(2).SetCellValue(25);

            HSSFFormulaEvaluator.EvaluateAllFormulaCells(wb);
            Assert.AreEqual(4.2 * 25, row.GetCell(3).NumericCellValue, 0.0001);


            if (OUTPUT_TEST_FILES)
            {
                // Save
                FileStream existing = File.Open(tmpDirName + "44636-existing.xls", FileMode.Open);
                existing.Seek(0, SeekOrigin.End);
                wb.Write(existing);
                existing.Close();
                Console.Error.WriteLine("Existing file for bug #44636 written to " + existing.ToString());
            }
            // Now, do a new file from scratch
            wb = new HSSFWorkbook();
            sheet = wb.CreateSheet();

            row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue(1.2);
            row.CreateCell(1).SetCellValue(4.2);

            row = sheet.CreateRow(1);
            row.CreateCell(0).CellFormula = ("SUM(A1:B1)");

            HSSFFormulaEvaluator.EvaluateAllFormulaCells(wb);
            Assert.AreEqual(5.4, row.GetCell(0).NumericCellValue, 0.0001);

            if (OUTPUT_TEST_FILES)
            {
                // Save
                FileStream scratch = File.Open(tmpDirName+"44636-scratch.xls",FileMode.Open);
                scratch.Seek(0, SeekOrigin.End);
                wb.Write(scratch);
                scratch.Close();
                Console.Error.WriteLine("New file for bug #44636 written to " + scratch.ToString());
            }
        }

        /**
         * Bug 44297: 32767+32768 is Evaluated to -1
         * Fix: IntPtg must operate with unsigned short. Reading signed short results in incorrect formula calculation
         * if a formula has values in the interval [Short.MAX_VALUE, (Short.MAX_VALUE+1)*2]
         *
         * @author Yegor Kozlov
         */
        [Test]
        public void Test44297()
        {

            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("44297.xls");

            IRow row;
            ICell cell;

            NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(0);

            HSSFFormulaEvaluator eva = new HSSFFormulaEvaluator(wb);

            row = sheet.GetRow(0);
            cell = row.GetCell(0);
            Assert.AreEqual("31+46", cell.CellFormula);
            Assert.AreEqual(77, eva.Evaluate(cell).NumberValue, 0);

            row = sheet.GetRow(1);
            cell = row.GetCell(0);
            Assert.AreEqual("30+53", cell.CellFormula);
            Assert.AreEqual(83, eva.Evaluate(cell).NumberValue, 0);

            row = sheet.GetRow(2);
            cell = row.GetCell(0);
            Assert.AreEqual("SUM(A1:A2)", cell.CellFormula);
            Assert.AreEqual(160, eva.Evaluate(cell).NumberValue, 0);

            row = sheet.GetRow(4);
            cell = row.GetCell(0);
            Assert.AreEqual("32767+32768", cell.CellFormula);
            Assert.AreEqual(65535, eva.Evaluate(cell).NumberValue, 0);

            row = sheet.GetRow(7);
            cell = row.GetCell(0);
            Assert.AreEqual("32744+42333", cell.CellFormula);
            Assert.AreEqual(75077, eva.Evaluate(cell).NumberValue, 0);

            row = sheet.GetRow(8);
            cell = row.GetCell(0);
            Assert.AreEqual("327680/32768", cell.CellFormula);
            Assert.AreEqual(10, eva.Evaluate(cell).NumberValue, 0);

            row = sheet.GetRow(9);
            cell = row.GetCell(0);
            Assert.AreEqual("32767+32769", cell.CellFormula);
            Assert.AreEqual(65536, eva.Evaluate(cell).NumberValue, 0);

            row = sheet.GetRow(10);
            cell = row.GetCell(0);
            Assert.AreEqual("35000+36000", cell.CellFormula);
            Assert.AreEqual(71000, eva.Evaluate(cell).NumberValue, 0);

            row = sheet.GetRow(11);
            cell = row.GetCell(0);
            Assert.AreEqual("-1000000-3000000", cell.CellFormula);
            Assert.AreEqual(-4000000, eva.Evaluate(cell).NumberValue, 0);
        }

        /**
         * Bug 44410: SUM(C:C) is valid in excel, and means a sum
         *  of all the rows in Column C
         *
         * @author Nick Burch
         */
        [Test]
        public void Test44410()
        {

            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("SingleLetterRanges.xls");

            NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(0);

            HSSFFormulaEvaluator eva = new HSSFFormulaEvaluator(wb);

            // =index(C:C,2,1) -> 2
            IRow rowIDX = sheet.GetRow(3);
            // =sum(C:C) -> 6
            IRow rowSUM = sheet.GetRow(4);
            // =sum(C:D) -> 66
            IRow rowSUM2D = sheet.GetRow(5);

            // Test the sum
            ICell cellSUM = rowSUM.GetCell(0);

            FormulaRecordAggregate frec = (FormulaRecordAggregate)((HSSFCell)cellSUM).CellValueRecord;
            Ptg[] ops = frec.FormulaRecord.ParsedExpression;
            Assert.AreEqual(2, ops.Length);
            Assert.AreEqual(typeof(AreaPtg), ops[0].GetType());
            Assert.AreEqual(typeof(FuncVarPtg), ops[1].GetType());

            // Actually stored as C1 to C65536
            // (last row is -1 === 65535)
            AreaPtg ptg = (AreaPtg)ops[0];
            Assert.AreEqual(2, ptg.FirstColumn);
            Assert.AreEqual(2, ptg.LastColumn);
            Assert.AreEqual(0, ptg.FirstRow);
            Assert.AreEqual(65535, ptg.LastRow);
            Assert.AreEqual("C:C", ptg.ToFormulaString());

            // Will show as C:C, but won't know how many
            // rows it covers as we don't have the sheet
            // to hand when turning the Ptgs into a string
            Assert.AreEqual("SUM(C:C)", cellSUM.CellFormula);

            // But the evaluator knows the sheet, so it
            // can do it properly
            Assert.AreEqual(6, eva.Evaluate(cellSUM).NumberValue, 0);

            // Test the index
            // Again, the formula string will be right but
            // lacking row count, Evaluated will be right
            ICell cellIDX = rowIDX.GetCell(0);
            Assert.AreEqual("INDEX(C:C,2,1)", cellIDX.CellFormula);
            Assert.AreEqual(2, eva.Evaluate(cellIDX).NumberValue, 0);

            // Across two colums
            ICell cellSUM2D = rowSUM2D.GetCell(0);
            Assert.AreEqual("SUM(C:D)", cellSUM2D.CellFormula);
            Assert.AreEqual(66, eva.Evaluate(cellSUM2D).NumberValue, 0);
        }

        /**
         * Tests that we can Evaluate boolean cells properly
         */
        [Test]
        public void TestEvaluateBooleanInCell_bug44508()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet();
            wb.SetSheetName(0, "Sheet1");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            cell.CellFormula = ("1=1");

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            try
            {
                fe.EvaluateInCell(cell);
            }
            catch (FormatException)
            {
                Assert.Fail("Identified bug 44508");
            }
            Assert.AreEqual(true, cell.BooleanCellValue);
        }
        [Test]
        public void TestClassCast_bug44861()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("44861.xls");

            // Check direct
            HSSFFormulaEvaluator.EvaluateAllFormulaCells(wb);

            // And via calls
            int numSheets = wb.NumberOfSheets;
            for (int i = 0; i < numSheets; i++)
            {
                NPOI.SS.UserModel.ISheet s = wb.GetSheetAt(i);
                HSSFFormulaEvaluator eval = new HSSFFormulaEvaluator(wb);

                for (IEnumerator rows = s.GetRowEnumerator(); rows.MoveNext(); )
                {
                    IRow r = (IRow)rows.Current;

                    for (IEnumerator cells = r.GetEnumerator(); cells.MoveNext(); )
                    {
                        ICell c = (ICell)cells.Current;
                        eval.EvaluateFormulaCell(c);
                    }
                }
            }
        }
        [Test]
        public void TestEvaluateInCellWithErrorCode_bug44950()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(1);
            ICell cell = row.CreateCell(0);
            cell.CellFormula = ("na()"); // this formula Evaluates to an Excel error code '#N/A'
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            try
            {
                fe.EvaluateInCell(cell);
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.StartsWith("Cannot get a error value from"))
                {
                    throw new AssertionException("Identified bug 44950 b");
                }
                throw;
            }
        }

        private class EvalListener : EvaluationListener
        {
            private int _countCacheHits;
            private int _countCacheMisses;

            public EvalListener()
            {
                _countCacheHits = 0;
                _countCacheMisses = 0;
            }
            public int GetCountCacheHits()
            {
                return _countCacheHits;
            }
            public int GetCountCacheMisses()
            {
                return _countCacheMisses;
            }

            public override void OnCacheHit(int sheetIndex, int srcRowNum, int srcColNum, ValueEval result)
            {
                _countCacheHits++;
            }
            public override void OnStartEvaluate(IEvaluationCell cell, ICacheEntry entry)
            {
                _countCacheMisses++;
            }
        }

        /**
         * The HSSFFormula evaluator performance benefits greatly from caching of intermediate cell values
         */
        [Test]
        public void TestSlowEvaluate45376()
        {
            /*
 * Note - to observe behaviour without caching, disable the call to
 * updateValue() from FormulaCellCacheEntry.updateFormulaResult().
 */

            // Firstly set up a sequence of formula cells where each depends on the  previous multiple
            // times.  Without caching, each subsequent cell take about 4 times longer to Evaluate.
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            for (int i = 1; i < 10; i++)
            {
                ICell cell = row.CreateCell(i);
                char prevCol = (char)('A' + i - 1);
                String prevCell = prevCol + "1";
                // this formula is inspired by the offending formula of the attachment for bug 45376
                String formula = "IF(DATE(YEAR(" + prevCell + "),MONTH(" + prevCell + ")+1,1)<=$D$3," +
                        "DATE(YEAR(" + prevCell + "),MONTH(" + prevCell + ")+1,1),NA())";
                cell.CellFormula = (formula);

            }
            row.CreateCell(0).SetCellValue(new DateTime(2000,1,1,0,0,0));

            // Choose cell A9, so that the Assert.Failing Test case doesn't take too long to execute.
            ICell cell1 = row.GetCell(8);
            EvalListener evalListener = new EvalListener();
            WorkbookEvaluator evaluator = WorkbookEvaluatorTestHelper.CreateEvaluator(wb, evalListener);
            ValueEval ve = evaluator.Evaluate(HSSFEvaluationTestHelper.WrapCell(cell1));
            int evalCount = evalListener.GetCountCacheMisses();
            if (evalCount > 10)
            {
                // Without caching, evaluating cell 'A9' takes 21845 evaluations which consumes
                // much time (~3 sec on Core 2 Duo 2.2GHz)
                Console.Error.WriteLine("Cell A9 took " + evalCount + " intermediate evaluations");
                throw new AssertionException("Identifed bug 45376 - Formula evaluator should cache values");
            }
            // With caching, the evaluationCount is 8 which is a big improvement
            // Note - these expected values may change if the WorkbookEvaluator is 
            // ever optimised to short circuit 'if' functions.
            Assert.AreEqual(8, evalCount);

            // The cache hits would be 24 if fully evaluating all arguments of the
            // "IF()" functions (Each of the 8 formulas has 4 refs to formula cells
            // which result in 1 cache miss and 3 cache hits). However with the
            // short-circuit-if optimisation, 2 of the cell refs get skipped
            // reducing this metric 8.
            Assert.AreEqual(8, evalListener.GetCountCacheHits());

            // confirm the evaluation result too
            Assert.AreEqual(ErrorEval.NA, ve);
        }
        [Test]
        public void TestDateWithNegativeParts_bug48528()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)wb.CreateSheet("Sheet1");
            HSSFRow row = (HSSFRow)sheet.CreateRow(1);
            HSSFCell cell = (HSSFCell)row.CreateCell(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

            // 5th Feb 2012 = 40944
            // 1st Feb 2012 = 40940
            // 5th Jan 2012 = 40913
            // 5th Dec 2011 = 40882
            // 5th Feb 2011 = 40579

            cell.CellFormula=("DATE(2012,2,1)");
            fe.NotifyUpdateCell(cell);
            Assert.AreEqual(40940.0, fe.Evaluate(cell).NumberValue);

            cell.CellFormula=("DATE(2012,2,1+4)");
            fe.NotifyUpdateCell(cell);
            Assert.AreEqual(40944.0, fe.Evaluate(cell).NumberValue);

            cell.CellFormula=("DATE(2012,2-1,1+4)");
            fe.NotifyUpdateCell(cell);
            Assert.AreEqual(40913.0, fe.Evaluate(cell).NumberValue);

            cell.CellFormula=("DATE(2012,2,1-27)");
            fe.NotifyUpdateCell(cell);
            Assert.AreEqual(40913.0, fe.Evaluate(cell).NumberValue);

            cell.CellFormula=("DATE(2012,2-2,1+4)");
            fe.NotifyUpdateCell(cell);
            Assert.AreEqual(40882.0, fe.Evaluate(cell).NumberValue);

            cell.CellFormula=("DATE(2012,2,1-58)");
            fe.NotifyUpdateCell(cell);
            Assert.AreEqual(40882.0, fe.Evaluate(cell).NumberValue);

            cell.CellFormula=("DATE(2012,2-12,1+4)");
            fe.NotifyUpdateCell(cell);
            Assert.AreEqual(40579.0, fe.Evaluate(cell).NumberValue);
        }
    }
}