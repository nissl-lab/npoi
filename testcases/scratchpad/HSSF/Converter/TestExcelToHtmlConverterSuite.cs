using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Converter;
using System.IO;

namespace TestCases.HSSF.Converter
{
    /// <summary>
    /// TestExcelToHtmlConverterSuite 的摘要说明
    /// </summary>
    [TestClass]
    public class TestExcelToHtmlConverterSuite
    {
        private static List<String> failingFiles = new List<string>();
        public TestExcelToHtmlConverterSuite()
        {
            //
            //TODO: 在此处添加构造函数逻辑
            //
        }

        private TestContext testContextInstance;

        /// <summary>
        ///获取或设置测试上下文，该上下文提供
        ///有关当前测试运行及其功能的信息。
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region 附加测试属性
        //
        // 编写测试时，还可使用以下附加属性:
        //
        // 在运行类中的第一个测试之前使用 ClassInitialize 运行代码
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // 在类中的所有测试都已运行之后使用 ClassCleanup 运行代码
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // 在运行每个测试之前，使用 TestInitialize 来运行代码
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // 在每个测试运行完之后，使用 TestCleanup 来运行代码
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
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
                excelToHtmlConverter.GetDocument().Save(Path.ChangeExtension(fileName, "html")); ;
        }
    }
}
