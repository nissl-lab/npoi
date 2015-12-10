using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using TestCases.HSSF;
using NPOI.SS.UserModel;
using NPOI.Util;
using System.IO;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.HSSF.UserModel;
using System.Globalization;

namespace NPOI.XSSF.UserModel
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
            //Workbook wbStreamingWritten = SXSSFITestDataProvider.instance.WriteOutAndReadBack(new SXSSFWorkbook(wb));
            //verifyBug54084Unicode(wbStreamingWritten);
        }

        private void verifyBug54084Unicode(XSSFWorkbook wb)
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
        }
        [Ignore]
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

        [Test]
        public void Test57236()
        {
            // Having very small numbers leads to different formatting, Excel uses the scientific notation, but POI leads to "0"

            /*
            DecimalFormat format = new DecimalFormat("#.##########", new DecimalFormatSymbols(Locale.Default));
            double d = 3.0E-104;
            Assert.AreEqual("3.0E-104", format.Format(d));
             */

            DataFormatter formatter = new DataFormatter(true);

            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57236.xlsx");
            for (int sheetNum = 0; sheetNum < wb.NumberOfSheets; sheetNum++)
            {
                ISheet sheet = wb.GetSheetAt(sheetNum);
                for (int rowNum = sheet.FirstRowNum; rowNum < sheet.LastRowNum; rowNum++)
                {
                    IRow row = sheet.GetRow(rowNum);
                    for (int cellNum = row.FirstCellNum; cellNum < row.LastCellNum; cellNum++)
                    {
                        ICell cell = row.GetCell(cellNum);
                        String fmtCellValue = formatter.FormatCellValue(cell);

                        //System.out.Println("Cell: " + fmtCellValue);
                        Assert.IsNotNull(fmtCellValue);
                        Assert.IsFalse(fmtCellValue.Equals("0"));
                    }
                }
            }
        }
    }
}

