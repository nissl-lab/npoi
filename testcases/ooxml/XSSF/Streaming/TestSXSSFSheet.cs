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

namespace NPOI.XSSF.streaming
{
    using System;


    using NPOI.SS;
    using NPOI.SS.UserModel;
    using NPOI.XSSF;
    using NPOI.XSSF.Streaming;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using TestCases.SS.UserModel;

    [TestFixture]
    public class TestSXSSFSheet : BaseTestSheet
    {

        public TestSXSSFSheet()
            : base(SXSSFITestDataProvider.instance)
        {
            ;
        }


        [TearDown]
        public void TearDown() {
            SXSSFITestDataProvider.instance.Cleanup();
        }

        /**
         * cloning of sheets is not supported in SXSSF
         */

        [Test]
        public void CloneSheet() {
            //thrown.Expect(typeof(Exception));
            //thrown.ExpectMessage("NotImplemented");
            base.TestCloneSheet();
        }


        [Test]
        public void CloneSheetMultipleTimes() {
            //thrown.Expect(typeof(Exception));
            //thrown.ExpectMessage("NotImplemented");
            base.TestCloneSheetMultipleTimes();
        }

        /**
         * Shifting rows is not supported in SXSSF
         */

        [Test]
        public void ShiftMerged() {
            //thrown.Expect(typeof(Exception));
            //thrown.ExpectMessage("NotImplemented");
            base.TestShiftMerged();
        }

        /**
         *  Bug 35084: cloning cells with formula
         *
         *  The test is disabled because cloning of sheets is not supported in SXSSF
         */

        [Test]
        public void bug35084() {
            //thrown.Expect(typeof(Exception));
            //thrown.ExpectMessage("NotImplemented");
            base.Test35084();
        }

        [Test]
        public void getCellComment() {
            // TODO: Reading cell comments via Sheet does not work currently as it tries 
            // to access the underlying sheet for this, but comments are stored as
            // properties on Cells...
        }


        [Test]
        public void defaultColumnStyle() {
            //TODO column styles are not yet supported by XSSF
        }

        [Test]
        public void OverrideFlushedRows() {
            IWorkbook wb = new SXSSFWorkbook(3);
            try {
                ISheet sheet = wb.CreateSheet();

                sheet.CreateRow(1);
                sheet.CreateRow(2);
                sheet.CreateRow(3);
                sheet.CreateRow(4);

                //thrown.Expect(typeof(Throwable));
                //thrown.ExpectMessage("Attempting to write a row[1] in the range [0,1] that is already written to disk.");
                sheet.CreateRow(1);
            } finally {
                wb.Close();
            }
        }

        [Test]
        public void OverrideRowsInTemplate() {
            XSSFWorkbook template = new XSSFWorkbook();
            template.CreateSheet().CreateRow(1);

            IWorkbook wb = new SXSSFWorkbook(template);
            try {
                ISheet sheet = wb.GetSheetAt(0);

                try {
                    sheet.CreateRow(1);
                    Assert.Fail("expected exception");
                } catch (Exception e) {
                    Assert.AreEqual("Attempting to write a row[1] in the range [0,1] that is already written to disk.", e.Message);
                }
                try {
                    sheet.CreateRow(0);
                    Assert.Fail("expected exception");
                } catch (Exception e) {
                    Assert.AreEqual("Attempting to write a row[0] in the range [0,1] that is already written to disk.", e.Message);
                }
                sheet.CreateRow(2);
            } finally {
                wb.Close();
                template.Close();
            }
        }

        [Test]
        public void CreateRowAfterLastRow() {
            createRowAfterLastRow(SpreadsheetVersion.EXCEL2007);
        }
    }
}
