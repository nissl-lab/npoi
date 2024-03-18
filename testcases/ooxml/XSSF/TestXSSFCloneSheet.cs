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

namespace TestCases.XSSF
{
    using System;
    using System.IO;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using TestCases.HSSF;
    using TestCases.SS.UserModel;

    [TestFixture]
    public class TestXSSFCloneSheet : BaseTestCloneSheet
    {
        public TestXSSFCloneSheet()
            : base(HSSFITestDataProvider.Instance)
        {
        }

        private static String OTHER_SHEET_NAME = "Another";
        private static String VALID_SHEET_NAME = "Sheet01";
        private XSSFWorkbook wb;

        [SetUp]
        public void SetUp() {
            wb = new XSSFWorkbook();
            wb.CreateSheet(VALID_SHEET_NAME);
        }

        [Test]
        public void TestCloneSheetIntStringValidName() {
            ISheet cloned = wb.CloneSheet(0, OTHER_SHEET_NAME);
            Assert.AreEqual(OTHER_SHEET_NAME, cloned.SheetName);
            Assert.AreEqual(2, wb.NumberOfSheets);
        }

        [Test]
        public void TestCloneSheetIntStringInvalidName() {
            try {
                wb.CloneSheet(0, VALID_SHEET_NAME);
                Assert.Fail("Should fail");
            } catch (ArgumentException) {
                // expected here
            }
            Assert.AreEqual(1, wb.NumberOfSheets);
        }

        [Test]
        public void Test60512()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("60512.xlsm");

            Assert.AreEqual(1, wb.NumberOfSheets);
            ISheet sheet = wb.CloneSheet(0);
            Assert.IsNotNull(sheet);
            Assert.AreEqual(2, wb.NumberOfSheets);


            IWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.IsNotNull(wbBack);
            wbBack.Close();

            wb.Close();
        }
    }

}