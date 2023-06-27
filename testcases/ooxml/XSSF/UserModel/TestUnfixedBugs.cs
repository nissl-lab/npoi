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

using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using TestCases.HSSF;

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    public class TestUnfixedBugs
    {
        [Test]
        public void TestBug54084Unicode()
        {
            // sample XLSX with the same text-contents as the text-file above
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("54084 - Greek - beyond BMP.xlsx");

            verifyBug54084Unicode(wb);

            //        OutputStream baos = new FileOutputStream("/tmp/test.xlsx");
            //        try {
            //            wb.Write(baos);
            //        } finally {
            //            baos.Close();
            //        }

            // now write the file and read it back in
            XSSFWorkbook wbWritten = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb);
            verifyBug54084Unicode(wbWritten);

            // finally also write it out via the streaming interface and verify that we still can read it back in
            SXSSFWorkbook swb = new SXSSFWorkbook(wb);
            IWorkbook wbStreamingWritten = SXSSFITestDataProvider.instance.WriteOutAndReadBack(swb);
            verifyBug54084Unicode(wbStreamingWritten);

            wbWritten.Close();
            swb.Close();
            wbStreamingWritten.Close();
            wb.Close();
        }

        private void verifyBug54084Unicode(IWorkbook wb)
        {
            // expected data is stored in UTF-8 in a text-file
            byte[] data = HSSFTestDataSamples.GetTestDataFileContent("54084 - Greek - beyond BMP.txt");
            String testData = Encoding.UTF8.GetString(data).Trim();

            ISheet sheet = wb.GetSheetAt(0);
            IRow row = sheet.GetRow(0);
            ICell cell = row.GetCell(0);

            String value = cell.StringCellValue;
            //Console.WriteLine(value);

            Assert.AreEqual(testData, value, "The data in the text-file should exactly match the data that we read from the workbook");
        }
        [Test]
        public void Test54071()
        {
            IWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("54071.xlsx");
            ISheet sheet = workbook.GetSheetAt(0);
            int rows = sheet.PhysicalNumberOfRows;
            Console.WriteLine(">> file rows is:" + (rows - 1) + " <<");
            IRow title = sheet.GetRow(0);

            DateTime? prev = null;
            for (int row = 1; row < rows; row++)
            {
                IRow rowObj = sheet.GetRow(row);
                for (int col = 0; col < 1; col++)
                {
                    String titleName = title.GetCell(col).ToString();
                    ICell cell = rowObj.GetCell(col);
                    if (titleName.StartsWith("time"))
                    {
                        // here the output will produce ...59 or ...58 for the rows, probably POI is
                        // doing some different rounding or some other small difference...
                        Console.WriteLine("==Time:" + cell.DateCellValue);
                        if (prev != null)
                        {
                            Assert.AreEqual(prev, cell.DateCellValue);
                            prev = cell.DateCellValue;
                        }
                    }
                }
            }
            workbook.Close();
        }
        [Ignore("test")]
        public void Test54071Simple()
        {
            double value1 = 41224.999988425923;
            double value2 = 41224.999988368058;

            int wholeDays1 = (int)Math.Floor(value1);
            int millisecondsInDay1 = (int)((value1 - wholeDays1) * DateUtil.DAY_MILLISECONDS + 0.5);

            int wholeDays2 = (int)Math.Floor(value2);
            int millisecondsInDay2 = (int)((value2 - wholeDays2) * DateUtil.DAY_MILLISECONDS + 0.5);

            Assert.AreEqual(wholeDays1, wholeDays2);
            // here we see that the time-value is 5 milliseconds apart, one is 86399000 and the other is 86398995, 
            // thus one is one second higher than the other
            Assert.AreEqual(millisecondsInDay1, millisecondsInDay2, "The time-values are 5 milliseconds apart");

            // when we do the calendar-stuff, there is a bool which determines if
            // the milliseconds are rounded or not, having this at "false" causes the 
            // second to be different here!
            int startYear = 1900;
            int dayAdjust = -1; // Excel thinks 2/29/1900 is a valid date, which it isn't
            //calendar1.Set(startYear, 0, wholeDays1 + dayAdjust, 0, 0, 0);
            //calendar1.Set(Calendar.MILLISECOND, millisecondsInDay1);
            DateTime calendar1 = new DateTime(startYear, 0, wholeDays1 + dayAdjust, 0, 0, 0);
            calendar1.AddMilliseconds(millisecondsInDay1);
            // this is the rounding part:
            //calendar1.Add(Calendar.MILLISECOND, 500);
            //calendar1.Clear(Calendar.MILLISECOND);
            calendar1.AddMilliseconds(500);
            calendar1.AddMilliseconds(-calendar1.Millisecond);

            DateTime calendar2 = new DateTime(startYear, 0, wholeDays2 + dayAdjust, 0, 0, 0);
            calendar2.AddMilliseconds(millisecondsInDay2);
            //calendar2.Set(startYear, 0, wholeDays2 + dayAdjust, 0, 0, 0);
            //calendar2.Set(Calendar.MILLISECOND, millisecondsInDay2);
            // this is the rounding part:
            //calendar2.Add(Calendar.MILLISECOND, 500);
            //calendar2.Clear(Calendar.MILLISECOND);
            calendar2.AddMilliseconds(500);
            calendar2.AddMilliseconds(-calendar2.Millisecond);

            // now the calendars are equal
            Assert.AreEqual(calendar1, calendar2);

            Assert.AreEqual(DateUtil.GetJavaDate(value1, false), DateUtil.GetJavaDate(value2, false));
        }


        // When this is fixed, the test case should go to BaseTestXCell with 
        // adjustments to use _testDataProvider to also verify this for XSSF
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestBug57294()
        {
            IWorkbook wb = SXSSFITestDataProvider.instance.CreateWorkbook();

            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            IRichTextString str = new XSSFRichTextString("Test rich text string");
            str.ApplyFont(2, 4, (short)0);
            Assert.AreEqual(3, str.NumFormattingRuns);
            cell.SetCellValue(str);

            IWorkbook wbBack = SXSSFITestDataProvider.instance.WriteOutAndReadBack(wb);
            wb.Close();

            // re-read after serializing and reading back
            ICell cellBack = wbBack.GetSheetAt(0).GetRow(0).GetCell(0);
            Assert.IsNotNull(cellBack);
            IRichTextString strBack = cellBack.RichStringCellValue;
            Assert.IsNotNull(strBack);
            Assert.AreEqual(3, strBack.NumFormattingRuns);
            Assert.AreEqual(0, strBack.GetIndexOfFormattingRun(0));
            Assert.AreEqual(2, strBack.GetIndexOfFormattingRun(1));
            Assert.AreEqual(4, strBack.GetIndexOfFormattingRun(2));

            wbBack.Close();
        }
        [Test]
        public void TestBug55752()
        {
            IWorkbook wb = new XSSFWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet("test");

                for (int i = 0; i < 4; i++)
                {
                    IRow row = sheet.CreateRow(i);
                    for (int j = 0; j < 2; j++)
                    {
                        ICell cell = row.CreateCell(j);
                        cell.CellStyle = (wb.CreateCellStyle());
                    }
                }

                // set content
                IRow row1 = sheet.GetRow(0);
                row1.GetCell(0).SetCellValue("AAA");
                IRow row2 = sheet.GetRow(1);
                row2.GetCell(0).SetCellValue("BBB");
                IRow row3 = sheet.GetRow(2);
                row3.GetCell(0).SetCellValue("CCC");
                IRow row4 = sheet.GetRow(3);
                row4.GetCell(0).SetCellValue("DDD");

                // merge cells
                CellRangeAddress range1 = new CellRangeAddress(0, 0, 0, 1);
                sheet.AddMergedRegion(range1);
                CellRangeAddress range2 = new CellRangeAddress(1, 1, 0, 1);
                sheet.AddMergedRegion(range2);
                CellRangeAddress range3 = new CellRangeAddress(2, 2, 0, 1);
                sheet.AddMergedRegion(range3);
                Assert.AreEqual(0, range3.FirstColumn);
                Assert.AreEqual(1, range3.LastColumn);
                Assert.AreEqual(2, range3.LastRow);
                CellRangeAddress range4 = new CellRangeAddress(3, 3, 0, 1);
                sheet.AddMergedRegion(range4);

                // set border
                RegionUtil.SetBorderBottom((int)BorderStyle.Thin, range1, sheet);

                row2.GetCell(0).CellStyle.BorderBottom = BorderStyle.Thin;
                row2.GetCell(1).CellStyle.BorderBottom = BorderStyle.Thin;
                ICell cell0 = CellUtil.GetCell(row3, 0);
                CellUtil.SetCellStyleProperty(cell0, CellUtil.BORDER_BOTTOM, BorderStyle.Thin);
                ICell cell1 = CellUtil.GetCell(row3, 1);
                CellUtil.SetCellStyleProperty(cell1, CellUtil.BORDER_BOTTOM, BorderStyle.Thin);
                RegionUtil.SetBorderBottom((int)BorderStyle.Thin, range4, sheet);

                // write to file
                Stream stream = new FileStream("55752.xlsx", FileMode.Create, FileAccess.ReadWrite);
                try
                {
                    wb.Write(stream, false);
                }
                finally
                {
                    stream.Close();
                }
            }
            finally
            {
                wb.Close();
            }
        }
        [Test]
        public void Test57423()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57423.xlsx");

            ISheet testSheet = wb.GetSheetAt(0);
            // row shift (negative or positive) causes corrupted output xlsx file when the shift value is bigger 
            // than the number of rows being shifted 
            // Excel 2010 on opening the output file says:
            // "Excel found unreadable content" and offers recovering the file by removing the unreadable content
            // This can be observed in cases like the following:
            // negative shift of 1 row by less than -1
            // negative shift of 2 rows by less than -2
            // positive shift of 1 row by 2 or more 
            // positive shift of 2 rows by 3 or more

            //testSheet.shiftRows(4, 5, -3);
            testSheet.ShiftRows(10, 10, 2);

            checkRows57423(testSheet);

            IWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            /*FileOutputStream stream = new FileOutputStream("C:\\temp\\57423.xlsx");
            try {
                wb.write(stream);
            } finally {
                stream.close();
            }*/
            wb.Close();

            checkRows57423(wbBack.GetSheetAt(0));

            wbBack.Close();
        }
        private void checkRows57423(ISheet testSheet)
        {
            checkRow57423(testSheet, 0, "0");
            checkRow57423(testSheet, 1, "1");
            checkRow57423(testSheet, 2, "2");
            checkRow57423(testSheet, 3, "3");
            checkRow57423(testSheet, 4, "4");
            checkRow57423(testSheet, 5, "5");
            checkRow57423(testSheet, 6, "6");
            checkRow57423(testSheet, 7, "7");
            checkRow57423(testSheet, 8, "8");
            checkRow57423(testSheet, 9, "9");

            Assert.IsNull(testSheet.GetRow(10),
                "Row number 10 should be gone after the shift");

            checkRow57423(testSheet, 11, "11");
            checkRow57423(testSheet, 12, "10");
            checkRow57423(testSheet, 13, "13");
            checkRow57423(testSheet, 14, "14");
            checkRow57423(testSheet, 15, "15");
            checkRow57423(testSheet, 16, "16");
            checkRow57423(testSheet, 17, "17");
            checkRow57423(testSheet, 18, "18");

            ByteArrayOutputStream stream = new ByteArrayOutputStream();
            try
            {
                ((XSSFSheet)testSheet).Write(stream);
            }
            finally
            {
                stream.Close();
            }

            // verify that the resulting XML has the rows in correct order as required by Excel
            String xml = Encoding.UTF8.GetString(stream.ToByteArray());
            int posR12 = xml.IndexOf("<row r=\"12\"");
            int posR13 = xml.IndexOf("<row r=\"13\"");

            // both need to be found
            Assert.IsTrue(posR12 != -1);
            Assert.IsTrue(posR13 != -1);

            Assert.IsTrue(posR12 < posR13,
                "Need to find row 12 before row 13 after the shifting, but had row 12 at " + posR12 + " and row 13 at " + posR13);
        }
        private void checkRow57423(ISheet testSheet, int rowNum, String contents)
        {
            IRow row = testSheet.GetRow(rowNum);
            Assert.IsNotNull(row, "Expecting row at rownum " + rowNum);

            CT_Row ctRow = ((XSSFRow)row).GetCTRow();
            Assert.AreEqual(rowNum + 1, ctRow.r);

            ICell cell = row.GetCell(0);
            Assert.IsNotNull(cell, "Expecting cell at rownum " + rowNum);
            //why concate ".0"? There is no ".0" in excel.
            Assert.AreEqual(contents, cell.ToString(), "Did not have expected contents at rownum " + rowNum);
            //Assert.AreEqual(contents + ".0", cell.ToString(), "Did not have expected contents at rownum " + rowNum);
        }

        [Test]
        public void test58325_one()
        {
            check58325(XSSFTestDataSamples.OpenSampleWorkbook("58325_lt.xlsx"), 1);
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void test58325_three()
        {
            check58325(XSSFTestDataSamples.OpenSampleWorkbook("58325_db.xlsx"), 3);
        }
        private void check58325(XSSFWorkbook wb, int expectedShapes)
        {
            XSSFSheet sheet = wb.GetSheet("MetasNM001") as XSSFSheet;
            Assert.IsNotNull(sheet);
            StringBuilder str = new StringBuilder();
            str.Append("sheet " + sheet.SheetName + " - ");
            XSSFDrawing drawing = sheet.GetDrawingPatriarch();
            //drawing = ((XSSFSheet)sheet).createDrawingPatriarch();
            List<XSSFShape> shapes = drawing.GetShapes();
            str.Append("drawing.Shapes.size() = " + shapes.Count);
            IEnumerator<XSSFShape> it = shapes.GetEnumerator();
            while (it.MoveNext())
            {
                XSSFShape shape = it.Current;
                str.Append(", " + shape.ToString());
                str.Append(", Col1:" + ((XSSFClientAnchor)shape.GetAnchor()).Col1);
                str.Append(", Col2:" + ((XSSFClientAnchor)shape.GetAnchor()).Col2);
                str.Append(", Row1:" + ((XSSFClientAnchor)shape.GetAnchor()).Row1);
                str.Append(", Row2:" + ((XSSFClientAnchor)shape.GetAnchor()).Row2);
            }

            Assert.AreEqual(expectedShapes, shapes.Count, 
                "Having shapes: " + str);
        }

    }
}

