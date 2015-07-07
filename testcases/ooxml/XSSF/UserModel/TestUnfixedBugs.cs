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

            String testData = Encoding.UTF8.GetString(HSSFTestDataSamples.GetTestDataFileContent("54084 - Greek - beyond BMP.txt")).Trim();

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
        
    }
}

