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
using System.Collections;
namespace TestCases.XSSF.IO
{

    [TestFixture]
    public class TestLoadSaveXSSF
    {
        private static POIDataSamples _ssSampels = POIDataSamples.GetSpreadSheetInstance();

        [Test]
        public void TestLoadSample()
        {
            XSSFWorkbook workbook = new XSSFWorkbook(_ssSampels.OpenResourceAsStream("sample.xlsx"));
            Assert.AreEqual(3, workbook.NumberOfSheets);
            Assert.AreEqual("Sheet1", workbook.GetSheetName(0));
            ISheet sheet = workbook.GetSheetAt(0);
            IRow row = sheet.GetRow(0);
            ICell cell = row.GetCell((short)1);
            Assert.IsNotNull(cell);
            Assert.AreEqual(111.0, cell.NumericCellValue, 0.0);
            cell = row.GetCell((short)0);
            Assert.AreEqual("Lorem", cell.RichStringCellValue.String);
        }

        // TODO filename string hard coded in XSSFWorkbook constructor in order to make ant Test-ooxml target be successful.
        [Test]
        public void TestLoadStyles()
        {
            XSSFWorkbook workbook = new XSSFWorkbook(_ssSampels.OpenResourceAsStream("styles.xlsx"));
            ISheet sheet = workbook.GetSheetAt(0);
            IRow row = sheet.GetRow(0);
            ICell cell = row.GetCell((short)0);
            ICellStyle style = cell.CellStyle;
            // assertNotNull(style);
        }

        // TODO filename string hard coded in XSSFWorkbook constructor in order to make ant Test-ooxml target be successful.
        [Test]
        public void TestLoadPictures()
        {
            XSSFWorkbook workbook = new XSSFWorkbook(_ssSampels.OpenResourceAsStream("picture.xlsx"));
            IList pictures = workbook.GetAllPictures();
            Assert.AreEqual(1, pictures.Count);
        }
    }
}


