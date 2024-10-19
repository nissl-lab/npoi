using NPOI.HSSF.UserModel;
using NPOI.SS.Converter;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.SS.Converter
{
    [TestFixture]
    public class TestExcelToHtmlConverterSuite
    {
        private static List<String> failingFiles = new List<string>();

        [Test]
        [Ignore("This will fail. The xls file may not be valid at all")]
        public void TestExcelToHtmlConverter()
        {
            string[] fileNames = POIDataSamples.GetSpreadSheetInstance().GetFiles("*.xls");
            List<string> toConverter = new List<string>();
            StringBuilder stringBuilder = new StringBuilder();
            foreach (string filename in fileNames)
            {
                if (!filename.EndsWith("clusterfuzz-testcase-minimized-POIHSSFFuzzer-6322470200934400.xls"))
                    toConverter.Add(filename);
                else
                    continue;
                try
                {
                    Test(filename);
                }
                catch (Exception ex)
                {
                    failingFiles.Add(filename);
                    stringBuilder.AppendLine(filename);
                    stringBuilder.AppendLine(ex.Source);
                    stringBuilder.AppendLine(ex.Message);
                    stringBuilder.AppendLine(ex.StackTrace);
                    stringBuilder.AppendLine("**************************************");
                }
            }

            string output = string.Empty;
            if (failingFiles.Count > 0)
            {
                output = Path.GetDirectoryName(failingFiles[0]) + "\\failxls.txt";
                using (StreamWriter sw = new StreamWriter(output, false))
                {
                    foreach (string file in failingFiles)
                    {
                        sw.WriteLine(file);
                    }
                    sw.WriteLine("**********************************************************");
                    sw.Write(stringBuilder.ToString());
                    sw.Close();
                }
            }
            Assert.IsTrue(failingFiles.Count == 0, "{0}({1}) files failed to convert to html. see " + output, failingFiles.Count, toConverter.Count);
        }
        private void Test(string fileName)
        {
            HSSFWorkbook workbook;
            workbook = ExcelToHtmlUtils.LoadXls(fileName);
            ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter();
            excelToHtmlConverter.ProcessWorkbook(workbook);
            excelToHtmlConverter.Document.Save(Path.ChangeExtension(fileName, "html")); ;
        }

        [Test]
        public void TestExcelToHtmlConverterWithBackground()
        {
           XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("background_color.xlsx");           

            ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter();
            excelToHtmlConverter.ProcessWorkbook(workbook);
            excelToHtmlConverter.Document.Save(Path.ChangeExtension("background_color", "html"));
        }
    }
}
