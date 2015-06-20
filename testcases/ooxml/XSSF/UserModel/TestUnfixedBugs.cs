using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using TestCases.HSSF;
using NPOI.SS.UserModel;

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
            //            wb.write(baos);
            //        } finally {
            //            baos.close();
            //        }

            // now write the file and read it back in
            XSSFWorkbook wbWritten = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb);
            verifyBug54084Unicode(wbWritten);

            // finally also write it out via the streaming interface and verify that we still can read it back in
            //Workbook wbStreamingWritten = SXSSFITestDataProvider.instance.writeOutAndReadBack(new SXSSFWorkbook(wb));
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
            //System.out.println(value);

            Assert.AreEqual(testData, value, "The data in the text-file should exactly match the data that we read from the workbook");
        }
    }
}
