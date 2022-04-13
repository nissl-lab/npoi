using NPOI;
using NPOI.OpenXml4Net.OPC;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System;
using System.Diagnostics;
using System.Xml;

namespace TestCases.OOXML
{
    [TestFixture]
    public class TestPOIXMLDocument2
    {
        [Test]
        public void TestMethod1()
        {
            Stopwatch watcher = new Stopwatch();
            watcher.Start();

            var filePath = @"D:\Documents\C04Tencent\02项目建设\20220331管家2022H1性能优化\资料\导入关键词30w物料模板.xlsx";
            XSSFWorkbook doc = new XSSFWorkbook(filePath);

            var sheet = doc.GetSheetAt(0);

            watcher.Stop();
            Console.WriteLine("GetExcelData by NOPI cost: {0}s", watcher.ElapsedMilliseconds / 1000);

            Test(doc, 4);
        }


        [Test]
        public void LoadXmlDocument()
        {
            Stopwatch watcher = new Stopwatch();
            watcher.Start();

            var filePath = @"D:\Documents\C04Tencent\02项目建设\20220331管家2022H1性能优化\资料\sheet2.xml";

            int i = 1;
            XmlDocument xml = new XmlDocument();
            xml.Load(filePath);
            XmlNodeList node = xml.SelectNodes("/worksheet/sheetData/row");
            foreach (XmlNode n in node)
            {
                XmlDocument x = new XmlDocument();
                XmlDeclaration dec = x.CreateXmlDeclaration("1.0", "utf-8", null);
                x.AppendChild(dec);
                XmlNode Product = x.ImportNode(n, true);
                x.AppendChild(Product);
                x.Save(@"E:\" + i + ".xml");
                i++;
            }

            watcher.Stop();
            Console.WriteLine("GetExcelData by NOPI cost: {0}s", watcher.ElapsedMilliseconds / 1000);

        }

        [Test]
        public void LoadXmlReader()
        {
            var filePath = @"D:\Documents\C04Tencent\02项目建设\20220331管家2022H1性能优化\资料\sheet2.xml";
            try
            {
                using (XmlReader reader = XmlReader.Create(filePath))
                {
                    string result;
                    while (reader.Read())
                    {
                        result = Trace(reader);
                        if (reader.NodeType == XmlNodeType.Element)
                        {
                            if (reader.Name == "row")
                            {
                            }
                            else if (reader.Name == "c")
                            {
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        private static string Trace(XmlReader reader)
        {
            string result = "";
            for (int count = 1; count <= reader.Depth; count++)
            {
                result += "===";
            }
            result += "=> " + reader.Name + "<br/>";
            System.Diagnostics.Trace.WriteLine(result);
            return result;
        }

        private void Test(POIXMLDocument doc, int expectedCount)
        {
            Assert.IsNotNull(doc.GetAllEmbedds());
            Assert.AreEqual(expectedCount, doc.GetAllEmbedds().Count);

            for (int i = 0; i < doc.GetAllEmbedds().Count; i++)
            {
                PackagePart pp = doc.GetAllEmbedds()[i];
                Assert.IsNotNull(pp);

                byte[] b = IOUtils.ToByteArray(pp.GetStream(System.IO.FileMode.Open));
                Assert.IsTrue(b.Length > 0);
            }
        }
    }
}
