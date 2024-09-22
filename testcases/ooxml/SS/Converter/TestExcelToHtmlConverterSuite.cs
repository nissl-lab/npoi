using NPOI.HSSF.UserModel;
using NPOI.SS.Converter;
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
       
        public void TestExcelToHtmlConverter()
        {
            string[] fileNames = POIDataSamples.GetSpreadSheetInstance().GetFiles("*.xls");
            List<string> toConverter = new List<string>();
            StringBuilder stringBuilder = new StringBuilder();
            foreach (string filename in fileNames)
            {
                if (filename.EndsWith("clusterfuzz-testcase-minimized-POIHSSFFuzzer-6322470200934400.xls"))
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
            //
            // TODO: 在此	添加测试逻辑
            //
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
            var fi = new FileInfo(@"..\..\..\..\test-data\spreadsheet\background_color.xlsx");
            string fileName = fi.FullName;

            XSSFWorkbook workbook;
            FileStream inputStream = File.Open(fileName, FileMode.Open);
            try
            {
                workbook = new XSSFWorkbook(inputStream);
            }
            finally
            {
                if (inputStream != null)
                    inputStream.Close();

                inputStream = null;                
            }

            ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter();
            excelToHtmlConverter.ProcessWorkbook(workbook);
            excelToHtmlConverter.Document.Save(Path.ChangeExtension(fileName, "html"));

        }
    }
}
