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
                    }
                }
            }
        }
        [Test]
        public void TestBug53798XLSX()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("53798_ShiftNegative_TMPL.xlsx");
            FileInfo xlsOutput = TempFile.CreateTempFile("testBug53798", ".xlsx");
            bug53798Work(wb, xlsOutput);
        }

        // Disabled because shift rows is not yet implemented for SXSSFWorkbook
        public void disabled_testBug53798XLSXStream()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("53798_ShiftNegative_TMPL.xlsx");
            FileInfo xlsOutput = TempFile.CreateTempFile("testBug53798", ".xlsx");
            //bug53798Work(new SXSSFWorkbook(wb), xlsOutput);
        }
        [Test]
        public void TestBug53798XLS()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("53798_ShiftNegative_TMPL.xls");
            FileInfo xlsOutput = TempFile.CreateTempFile("testBug53798", ".xls");
            bug53798Work(wb, xlsOutput);
        }

        private void bug53798Work(IWorkbook wb, FileInfo xlsOutput)
        {
            ISheet TestSheet = wb.GetSheetAt(0);

            TestSheet.ShiftRows(2, 2, 1);

            SaveAndReloadReport(wb, xlsOutput);

            // 1) corrupted xlsx (unreadable data in the first row of a Shifted group) already comes about
            // when Shifted by less than -1 negative amount (try -2)
            TestSheet.ShiftRows(3, 3, -1);

            SaveAndReloadReport(wb, xlsOutput);

            TestSheet.ShiftRows(2, 2, 1);

            SaveAndReloadReport(wb, xlsOutput);

            IRow newRow = null;
            ICell newCell = null;
            // 2) attempt to create a new row IN PLACE of a Removed row by a negative shift causes corrupted
            // xlsx file with  unreadable data in the negative Shifted row.
            // NOTE it's ok to create any other row.
            newRow = TestSheet.CreateRow(3);

            SaveAndReloadReport(wb, xlsOutput);

            newCell = newRow.CreateCell(0);

            SaveAndReloadReport(wb, xlsOutput);

            newCell.SetCellValue("new Cell in row " + newRow.RowNum);

            SaveAndReloadReport(wb, xlsOutput);

            // 3) once a negative shift has been made any attempt to shift another group of rows
            // (note: outside of previously negative Shifted rows) by a POSITIVE amount causes POI exception:
            // org.apache.xmlbeans.impl.values.XmlValueDisconnectedException.
            // NOTE: another negative shift on another group of rows is successful, provided no new rows in
            // place of previously Shifted rows were attempted to be Created as explained above.
            TestSheet.ShiftRows(6, 7, 1);	// -- CHANGE the shift to positive once the behaviour of
            // the above has been Tested

            SaveAndReloadReport(wb, xlsOutput);
        }

        private void SaveAndReloadReport(IWorkbook wb, FileInfo outFile)
        {
            // run some method on the font to verify if it is "disconnected" already
            //for(short i = 0;i < 256;i++)
            {
                IFont font = wb.GetFontAt((short)0);
                if (font is XSSFFont)
                {
                    XSSFFont xfont = (XSSFFont)wb.GetFontAt((short)0);
                    CT_Font ctFont = (CT_Font)xfont.GetCTFont();
                    Assert.AreEqual(0, ctFont.sizeOfBArray());
                }
            }

            FileStream fileOutStream = new FileStream(outFile.FullName, FileMode.Truncate, FileAccess.ReadWrite);
            wb.Write(fileOutStream);
            fileOutStream.Close();
            //Console.WriteLine("File \""+outFile.GetName()+"\" has been saved successfully");

            Stream is1 = new FileStream(outFile.FullName, FileMode.Truncate, FileAccess.ReadWrite);
            try
            {
                IWorkbook newWB;
                if (wb is XSSFWorkbook)
                {
                    newWB = new XSSFWorkbook(is1);
                }
                else if (wb is HSSFWorkbook)
                {
                    newWB = new HSSFWorkbook(is1);
                    //} else if(wb is SXSSFWorkbook) {
                    //    newWB = new SXSSFWorkbook(new XSSFWorkbook(is1));
                }
                else
                {
                    throw new InvalidOperationException("Unknown workbook: " + wb);
                }
                Assert.IsNotNull(newWB.GetSheet("test"));
            }
            finally
            {
                is1.Close();
            }
        }
    }
}

