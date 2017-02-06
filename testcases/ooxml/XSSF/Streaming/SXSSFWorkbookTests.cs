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
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using NUnit.Framework;

namespace NPOI.OOXML.Testcases.XSSF.Streaming
{
    [TestFixture]
    class SXSSFWorkbookTests
    {
        private SXSSFWorkbook _objectToTest { get; set; }

        [Test]
        public void
            CallingEmptyConstructorShouldInstanstiateNewXssfWorkbookDefaultRowAccessWindowSizeCompressTempFilesAsFalseAndUseSharedStringsTableFalse
            ()
        {
            _objectToTest = new SXSSFWorkbook();
            Assert.AreEqual(100, _objectToTest.RandomAccessWindowSize);
            Assert.NotNull(_objectToTest.XssfWorkbook);
        }

        [Test]
        public void
            CallingConstructorWithNullWorkbookShouldInstanstiateNewXssfWorkbookDefaultRowAccessWindowSizeCompressTempFilesAsFalseAndUseSharedStringsTableFalse
            ()
        {
            _objectToTest = new SXSSFWorkbook(null);
            Assert.AreEqual(100, _objectToTest.RandomAccessWindowSize);
            Assert.NotNull(_objectToTest.XssfWorkbook);
        }

        [Test]
        public void
    CallingConstructorWithExistingWorkbookShouldInstanstiateNewXssfWorkbookDefaultRowAccessWindowSizeCompressTempFilesAsFalseAndUseSharedStringsTableFalse
    ()
        {
            var wb = new XSSFWorkbook();
            var name = wb.CreateName();

            name.NameName = "test";
            var sheet = wb.CreateSheet("test1");

            _objectToTest = new SXSSFWorkbook(wb);
            Assert.AreEqual(100, _objectToTest.RandomAccessWindowSize);
            Assert.AreEqual("test", _objectToTest.XssfWorkbook.GetName("test").NameName);
            Assert.AreEqual(1, _objectToTest.NumberOfSheets);
        }

        [Test]
        public void IfCompressTmpFilesIsSetToTrueShouldReturnGZIPSheetDataWriter()
        {
            _objectToTest = new SXSSFWorkbook(null, 100, true);
            var result = _objectToTest.CreateSheetDataWriter();

            Assert.IsTrue(result is GZIPSheetDataWriter);

        }

        [Test]
        public void IfCompressTmpFilesIsSetToFalseShouldReturnSheetDataWriter()
        {
            _objectToTest = new SXSSFWorkbook();
            var result = _objectToTest.CreateSheetDataWriter();

            Assert.IsTrue(result is SheetDataWriter);

        }

        [Test]
        public void IfSettingSheetOrderShouldSetSheetOrderOfXssfWorkbook()
        {
            _objectToTest = new SXSSFWorkbook();
            _objectToTest.CreateSheet("test1");
            _objectToTest.CreateSheet("test2");

            _objectToTest.SetSheetOrder("test2", 0);

            Assert.AreEqual("test2", _objectToTest.GetSheetName(0));
            Assert.AreEqual("test1", _objectToTest.GetSheetName(1));
        }



    }
}
