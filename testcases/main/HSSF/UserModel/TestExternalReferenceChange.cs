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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace TestCases.HSSF.UserModel
{
    [TestFixture]
    public class TestExternalReferenceChange
    {

        private static String MAIN_WORKBOOK_FILENAME = "52575_main.xls";
        private static String SOURCE_DUMMY_WORKBOOK_FILENAME = "source_dummy.xls";
        private static String SOURCE_WORKBOOK_FILENAME = "52575_source.xls";

        private HSSFWorkbook mainWorkbook;
        private HSSFWorkbook sourceWorkbook;
        [SetUp]
        protected void SetUp()
        {
            mainWorkbook = HSSFTestDataSamples.OpenSampleWorkbook(MAIN_WORKBOOK_FILENAME);
            sourceWorkbook = HSSFTestDataSamples.OpenSampleWorkbook(SOURCE_WORKBOOK_FILENAME);

            Assert.IsNotNull(mainWorkbook);
            Assert.IsNotNull(sourceWorkbook);
        }
        [Test]
        public void TestDummyToSource()
        {
            bool changed = mainWorkbook.ChangeExternalReference("DOESNOTEXIST", SOURCE_WORKBOOK_FILENAME);
            Assert.IsFalse(changed);

            changed = mainWorkbook.ChangeExternalReference(SOURCE_DUMMY_WORKBOOK_FILENAME, SOURCE_WORKBOOK_FILENAME);
            Assert.IsTrue(changed);

            HSSFSheet lSheet = (HSSFSheet)mainWorkbook.GetSheetAt(0);
            HSSFCell lA1Cell = (HSSFCell)lSheet.GetRow(0).GetCell(0);

            Assert.AreEqual(CellType.Formula, lA1Cell.CellType);

            HSSFFormulaEvaluator lMainWorkbookEvaluator = new HSSFFormulaEvaluator(mainWorkbook);
            HSSFFormulaEvaluator lSourceEvaluator = new HSSFFormulaEvaluator(sourceWorkbook);
            HSSFFormulaEvaluator.SetupEnvironment(
                    new String[] { MAIN_WORKBOOK_FILENAME, SOURCE_WORKBOOK_FILENAME },
                    new HSSFFormulaEvaluator[] { lMainWorkbookEvaluator, lSourceEvaluator });

            Assert.AreEqual(CellType.Numeric, lMainWorkbookEvaluator.EvaluateFormulaCell(lA1Cell));

            Assert.AreEqual(20.0d, lA1Cell.NumericCellValue, 0.00001d);

        }

    }
}
