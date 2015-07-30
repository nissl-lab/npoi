using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.HSSF.UserModel;
using NUnit.Framework;
using TestCases;

namespace NPOI.SS.Converter
{
    [TestFixture]
    public class TestExcelToHtmlConverterSuite
    {
        private static List<String> failingFiles = new List<string>();

        [Test]
        [Ignore]
        public void TestExcelToHtmlConverter()
        {
            string[] fileNames = POIDataSamples.GetSpreadSheetInstance().GetFiles("*.xls");
            List<string> toConverter = new List<string>();
            StringBuilder stringBuilder = new StringBuilder();
            foreach (string filename in fileNames)
            {
                if (filename.EndsWith(".xls"))
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
    }
}
