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

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System.IO;

namespace NPOI.OOXML.Tests.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFSheetLazyLoad
    {
        /// <summary>
        /// Create a simple in-memory xlsx workbook for testing
        /// </summary>
        private static byte[] CreateSimpleXlsx()
        {
            using var wb = new XSSFWorkbook();
            var sheet = wb.CreateSheet("TestSheet");
            var row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("Hello");
            row.CreateCell(1).SetCellValue(42.0);
            sheet.CreateRow(1).CreateCell(0).SetCellValue("World");
            using var ms = new MemoryStream();
            wb.Write(ms);
            return ms.ToArray();
        }

        [Test]
        public void OpeningWorkbookDoesNotParseSheet()
        {
            byte[] xlsxBytes = CreateSimpleXlsx();
            using var ms = new MemoryStream(xlsxBytes);
            using var wb = new XSSFWorkbook(ms);

            // Getting a sheet should NOT trigger parsing
            XSSFSheet sheet = (XSSFSheet)wb.GetSheetAt(0);

            // _parseCount should be 0 - sheet.xml not parsed yet
            Assert.That(sheet._parseCount, Is.EqualTo(0),
                "Sheet XML should not be parsed merely by opening workbook or calling GetSheetAt");
        }

        [Test]
        public void GetSheetAtDoesNotParseSheet()
        {
            byte[] xlsxBytes = CreateSimpleXlsx();
            using var ms = new MemoryStream(xlsxBytes);
            using var wb = new XSSFWorkbook(ms);

            XSSFSheet sheet = (XSSFSheet)wb.GetSheetAt(0);
            Assert.That(sheet._parseCount, Is.EqualTo(0),
                "GetSheetAt() alone should not trigger sheet XML parsing");
        }

        [Test]
        public void GetSheetByNameDoesNotParseSheet()
        {
            byte[] xlsxBytes = CreateSimpleXlsx();
            using var ms = new MemoryStream(xlsxBytes);
            using var wb = new XSSFWorkbook(ms);

            XSSFSheet sheet = (XSSFSheet)wb.GetSheet("TestSheet");
            Assert.That(sheet._parseCount, Is.EqualTo(0),
                "GetSheet(name) alone should not trigger sheet XML parsing");
        }

        [Test]
        public void AccessingLastRowNumTriggersParseExactlyOnce()
        {
            byte[] xlsxBytes = CreateSimpleXlsx();
            using var ms = new MemoryStream(xlsxBytes);
            using var wb = new XSSFWorkbook(ms);

            XSSFSheet sheet = (XSSFSheet)wb.GetSheetAt(0);
            Assert.That(sheet._parseCount, Is.EqualTo(0), "Before access: no parse");

            // First access triggers parse
            int lastRow = sheet.LastRowNum;
            Assert.That(sheet._parseCount, Is.EqualTo(1), "After LastRowNum: exactly one parse");

            // Subsequent access does NOT re-parse
            lastRow = sheet.LastRowNum;
            Assert.That(sheet._parseCount, Is.EqualTo(1), "After second LastRowNum: still exactly one parse");
        }

        [Test]
        public void AccessingGetRowTriggersParseExactlyOnce()
        {
            byte[] xlsxBytes = CreateSimpleXlsx();
            using var ms = new MemoryStream(xlsxBytes);
            using var wb = new XSSFWorkbook(ms);

            XSSFSheet sheet = (XSSFSheet)wb.GetSheetAt(0);
            Assert.That(sheet._parseCount, Is.EqualTo(0), "Before access: no parse");

            IRow row = sheet.GetRow(0);
            Assert.That(sheet._parseCount, Is.EqualTo(1), "After GetRow: exactly one parse");
            Assert.That(row, Is.Not.Null, "Row 0 should exist");

            row = sheet.GetRow(1);
            Assert.That(sheet._parseCount, Is.EqualTo(1), "After second GetRow: still one parse");
        }

        [Test]
        public void IteratingRowsTriggersParseExactlyOnce()
        {
            byte[] xlsxBytes = CreateSimpleXlsx();
            using var ms = new MemoryStream(xlsxBytes);
            using var wb = new XSSFWorkbook(ms);

            XSSFSheet sheet = (XSSFSheet)wb.GetSheetAt(0);
            Assert.That(sheet._parseCount, Is.EqualTo(0), "Before iteration: no parse");

            int rowCount = 0;
            foreach (IRow row in sheet)
            {
                rowCount++;
            }

            Assert.That(sheet._parseCount, Is.EqualTo(1), "After iteration: exactly one parse");
            Assert.That(rowCount, Is.EqualTo(2), "Should have 2 rows");
        }

        [Test]
        public void NewSheetIsMarkedLoadedFromStart()
        {
            using var wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("NewSheet");

            // A newly created (not loaded from file) sheet is already "loaded"
            Assert.That(sheet._parseCount, Is.EqualTo(0),
                "New sheet was not parsed from file, parseCount should be 0");

            // But it should work immediately without parsing
            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("test");
            Assert.That(sheet.LastRowNum, Is.EqualTo(0), "New sheet should have data");
            Assert.That(sheet._parseCount, Is.EqualTo(0), "No parsing needed for new sheet");
        }

        [Test]
        public void LazyLoadedSheetDataIsCorrect()
        {
            byte[] xlsxBytes = CreateSimpleXlsx();
            using var ms = new MemoryStream(xlsxBytes);
            using var wb = new XSSFWorkbook(ms);

            XSSFSheet sheet = (XSSFSheet)wb.GetSheetAt(0);

            // Access data
            IRow row0 = sheet.GetRow(0);
            Assert.That(row0, Is.Not.Null);
            Assert.That(row0.GetCell(0).StringCellValue, Is.EqualTo("Hello"));
            Assert.That(row0.GetCell(1).NumericCellValue, Is.EqualTo(42.0));

            IRow row1 = sheet.GetRow(1);
            Assert.That(row1, Is.Not.Null);
            Assert.That(row1.GetCell(0).StringCellValue, Is.EqualTo("World"));

            Assert.That(sheet.LastRowNum, Is.EqualTo(1));
            Assert.That(sheet._parseCount, Is.EqualTo(1), "Parsed only once despite multiple accesses");
        }

        [Test]
        public void MultipleSheetsSomeAccessedSomeNot()
        {
            byte[] xlsxBytes;
            using (var wb = new XSSFWorkbook())
            {
                for (int i = 0; i < 3; i++)
                {
                    var s = wb.CreateSheet($"Sheet{i}");
                    s.CreateRow(0).CreateCell(0).SetCellValue(i);
                }
                using var ms2 = new MemoryStream();
                wb.Write(ms2);
                xlsxBytes = ms2.ToArray();
            }

            using var ms = new MemoryStream(xlsxBytes);
            using var wb2 = new XSSFWorkbook(ms);

            var sheet0 = (XSSFSheet)wb2.GetSheetAt(0);
            var sheet1 = (XSSFSheet)wb2.GetSheetAt(1);
            var sheet2 = (XSSFSheet)wb2.GetSheetAt(2);

            // None parsed yet
            Assert.That(sheet0._parseCount, Is.EqualTo(0));
            Assert.That(sheet1._parseCount, Is.EqualTo(0));
            Assert.That(sheet2._parseCount, Is.EqualTo(0));

            // Access only sheet1
            int val = (int)sheet1.GetRow(0).GetCell(0).NumericCellValue;
            Assert.That(val, Is.EqualTo(1));

            // Only sheet1 was parsed
            Assert.That(sheet0._parseCount, Is.EqualTo(0), "sheet0 was not accessed, should not be parsed");
            Assert.That(sheet1._parseCount, Is.EqualTo(1), "sheet1 was accessed, should be parsed once");
            Assert.That(sheet2._parseCount, Is.EqualTo(0), "sheet2 was not accessed, should not be parsed");
        }
    }
}
