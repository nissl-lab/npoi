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

        [Test]
        public void Test57165()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57171_57163_57165.xlsx");
            try
            {
                RemoveAllSheetsBut(3, wb);
                wb.CloneSheet(0); // Throws exception here
                wb.SetSheetName(1, "New Sheet");
                //saveWorkbook(wb, fileName);

                XSSFWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
                try
                {

                }
                finally
                {
                    wbBack.Close();
                }
            }
            finally
            {
                wb.Close();
            }
        }

        private static void RemoveAllSheetsBut(int sheetIndex, IWorkbook wb)
        {
            int sheetNb = wb.NumberOfSheets;
            // Move this sheet at the first position
            wb.SetSheetOrder(wb.GetSheetName(sheetIndex), 0);
            for (int sn = sheetNb - 1; sn > 0; sn--)
            {
                wb.RemoveSheetAt(sn);
            }
        }

    }
}

