/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
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

using System;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util;
using NUnit.Framework;
using TestCases.HSSF;
using System.Collections.Generic;

namespace TestCases.SS.Formula
{
    [TestFixture]
    public class TestMissingWorkbook
    {
        protected String MAIN_WORKBOOK_FILENAME = "52575_main.xls";
        protected String SOURCE_DUMMY_WORKBOOK_FILENAME = "source_dummy.xls";
        protected String SOURCE_WORKBOOK_FILENAME = "52575_source.xls";

        protected IWorkbook mainWorkbook;
        protected IWorkbook sourceWorkbook;
        public TestMissingWorkbook()
            : this("52575_main.xls", "source_dummy.xls", "52575_source.xls")
        {
            ;
        }
        protected TestMissingWorkbook(String MAIN_WORKBOOK_FILENAME,
                String SOURCE_DUMMY_WORKBOOK_FILENAME, String SOURCE_WORKBOOK_FILENAME)
        {
            this.MAIN_WORKBOOK_FILENAME = MAIN_WORKBOOK_FILENAME;
            this.SOURCE_DUMMY_WORKBOOK_FILENAME = SOURCE_DUMMY_WORKBOOK_FILENAME;
            this.SOURCE_WORKBOOK_FILENAME = SOURCE_WORKBOOK_FILENAME;
        }
        [SetUp]
        protected virtual void SetUp()
        {
            mainWorkbook = HSSFTestDataSamples.OpenSampleWorkbook(MAIN_WORKBOOK_FILENAME);
            sourceWorkbook = HSSFTestDataSamples.OpenSampleWorkbook(SOURCE_WORKBOOK_FILENAME);

            Assert.IsNotNull(mainWorkbook);
            Assert.IsNotNull(sourceWorkbook);
        }

        [Test]
        public void TestMissingWorkbookMissing()
        {
            IFormulaEvaluator evaluator = mainWorkbook.GetCreationHelper().CreateFormulaEvaluator();

            ISheet lSheet = mainWorkbook.GetSheetAt(0);
            IRow lARow = lSheet.GetRow(0);
            ICell lA1Cell = lARow.GetCell(0);

            Assert.AreEqual(CellType.Formula, lA1Cell.CellType);
            try
            {
                evaluator.EvaluateFormulaCell(lA1Cell);
                Assert.Fail("Missing external workbook reference exception expected!");
            }
            catch (RuntimeException re)
            {
                Assert.IsTrue(re.Message.IndexOf(SOURCE_DUMMY_WORKBOOK_FILENAME) != -1, "Unexpected exception: " + re);
            }
        }

        [Test]
        public void TestMissingWorkbookMissingOverride()
        {
            ISheet lSheet = mainWorkbook.GetSheetAt(0);
            ICell lA1Cell = lSheet.GetRow(0).GetCell(0);
            ICell lB1Cell = lSheet.GetRow(1).GetCell(0);
            ICell lC1Cell = lSheet.GetRow(2).GetCell(0);

            Assert.AreEqual(CellType.Formula, lA1Cell.CellType);
            Assert.AreEqual(CellType.Formula, lB1Cell.CellType);
            Assert.AreEqual(CellType.Formula, lC1Cell.CellType);

            // Check cached values
            Assert.AreEqual(10.0d, lA1Cell.NumericCellValue, 0.00001d);
            Assert.AreEqual("POI rocks!", lB1Cell.StringCellValue);
            Assert.AreEqual(true, lC1Cell.BooleanCellValue);

            // Evaluate
            IFormulaEvaluator evaluator = mainWorkbook.GetCreationHelper().CreateFormulaEvaluator();
            evaluator.IgnoreMissingWorkbooks = (true);

            Assert.AreEqual(CellType.Numeric, evaluator.EvaluateFormulaCell(lA1Cell));
            Assert.AreEqual(CellType.String, evaluator.EvaluateFormulaCell(lB1Cell));
            Assert.AreEqual(CellType.Boolean, evaluator.EvaluateFormulaCell(lC1Cell));

            Assert.AreEqual(10.0d, lA1Cell.NumericCellValue, 0.00001d);
            Assert.AreEqual("POI rocks!", lB1Cell.StringCellValue);
            Assert.AreEqual(true, lC1Cell.BooleanCellValue);
        }

        [Test]
        public void TestExistingWorkbook()
        {
            ISheet lSheet = mainWorkbook.GetSheetAt(0);
            ICell lA1Cell = lSheet.GetRow(0).GetCell(0);
            ICell lB1Cell = lSheet.GetRow(1).GetCell(0);
            ICell lC1Cell = lSheet.GetRow(2).GetCell(0);

            Assert.AreEqual(CellType.Formula, lA1Cell.CellType);
            Assert.AreEqual(CellType.Formula, lB1Cell.CellType);
            Assert.AreEqual(CellType.Formula, lC1Cell.CellType);

            IFormulaEvaluator lMainWorkbookEvaluator = mainWorkbook.GetCreationHelper().CreateFormulaEvaluator();
            IFormulaEvaluator lSourceEvaluator = sourceWorkbook.GetCreationHelper().CreateFormulaEvaluator();
            Dictionary<String, IFormulaEvaluator> workbooks = new Dictionary<String, IFormulaEvaluator>();
            workbooks.Add(MAIN_WORKBOOK_FILENAME, lMainWorkbookEvaluator);
            workbooks.Add(SOURCE_DUMMY_WORKBOOK_FILENAME, lSourceEvaluator);
            lMainWorkbookEvaluator.SetupReferencedWorkbooks(workbooks);

            Assert.AreEqual(CellType.Numeric, lMainWorkbookEvaluator.EvaluateFormulaCell(lA1Cell));
            Assert.AreEqual(CellType.String, lMainWorkbookEvaluator.EvaluateFormulaCell(lB1Cell));
            Assert.AreEqual(CellType.Boolean, lMainWorkbookEvaluator.EvaluateFormulaCell(lC1Cell));

            Assert.AreEqual(20.0d, lA1Cell.NumericCellValue, 0.00001d);
            Assert.AreEqual("Apache rocks!", lB1Cell.StringCellValue);
            Assert.AreEqual(false, lC1Cell.BooleanCellValue);
        }

    }
}