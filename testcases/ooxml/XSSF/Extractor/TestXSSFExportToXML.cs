/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using NPOI;
using NPOI.XSSF;
using NPOI.XSSF.Extractor;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace TestCases.XSSF.Extractor
{

    /**
     * @author Roberto Manicardi
     */
    [TestFixture]
    public class TestXSSFExportToXML
    {

        [Test]
        public void TestExportToXML()
        {

            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("CustomXMLMappings.xlsx");

            foreach (POIXMLDocumentPart p in wb.GetRelations())
            {

                if (!(p is MapInfo))
                {
                    continue;
                }
                MapInfo mapInfo = (MapInfo)p;

                XSSFMap map = mapInfo.GetXSSFMapById(1);
                XSSFExportToXml exporter = new XSSFExportToXml(map);
                MemoryStream os = new MemoryStream();
                exporter.ExportToXML(os, true);
                os.Position = 0;
                XmlDocument xml = new XmlDocument();
                xml.Load(os);
                Assert.IsNotNull(xml);
                Assert.IsTrue(xml.OuterXml != "");


                String docente = xml.SelectSingleNode("//DOCENTE").InnerText.Trim();
                String nome = xml.SelectSingleNode("//NOME").InnerText.Trim();
                String tutor = xml.SelectSingleNode("//TUTOR").InnerText.Trim();
                String cdl = xml.SelectSingleNode("//CDL").InnerText.Trim();
                String durata = xml.SelectSingleNode("//DURATA").InnerText.Trim();
                String argomento = xml.SelectSingleNode("//ARGOMENTO").InnerText.Trim();
                String progetto = xml.SelectSingleNode("//PROGETTO").InnerText.Trim();
                String crediti = xml.SelectSingleNode("//CREDITI").InnerText.Trim();

                Assert.AreEqual("ro", docente);
                Assert.AreEqual("ro", nome);
                Assert.AreEqual("ds", tutor);
                Assert.AreEqual("gs", cdl);
                Assert.AreEqual("g", durata);
                Assert.AreEqual("gvvv", argomento);
                Assert.AreEqual("aaaa", progetto);
                Assert.AreEqual("aa", crediti);
            }
        }
        [Test]
        public void TestExportToXMLInverseOrder()
        {

            XSSFWorkbook wb = XSSFTestDataSamples
                    .OpenSampleWorkbook("CustomXmlMappings-inverse-order.xlsx");

            MapInfo mapInfo = null;

            foreach (POIXMLDocumentPart p in wb.GetRelations())
            {

                if (!(p is MapInfo))
                {
                    continue;
                }
                mapInfo = (MapInfo)p;

                XSSFMap map = mapInfo.GetXSSFMapById(1);
                XSSFExportToXml exporter = new XSSFExportToXml(map);
                MemoryStream os = new MemoryStream();
                exporter.ExportToXML(os, true);
                os.Position = 0;
                XmlDocument xml = new XmlDocument();
                xml.Load(os);
                Assert.IsNotNull(xml);
                Assert.IsTrue(xml.OuterXml != "");

                String docente = xml.SelectSingleNode("//DOCENTE").InnerText.Trim();
                String nome = xml.SelectSingleNode("//NOME").InnerText.Trim();
                String tutor = xml.SelectSingleNode("//TUTOR").InnerText.Trim();
                String cdl = xml.SelectSingleNode("//CDL").InnerText.Trim();
                String durata = xml.SelectSingleNode("//DURATA").InnerText.Trim();
                String argomento = xml.SelectSingleNode("//ARGOMENTO").InnerText.Trim();
                String progetto = xml.SelectSingleNode("//PROGETTO").InnerText.Trim();
                String crediti = xml.SelectSingleNode("//CREDITI").InnerText.Trim();

                Assert.AreEqual("aa", nome);
                Assert.AreEqual("aaaa", docente);
                Assert.AreEqual("gvvv", tutor);
                Assert.AreEqual("g", cdl);
                Assert.AreEqual("gs", durata);
                Assert.AreEqual("ds", argomento);
                Assert.AreEqual("ro", progetto);
                Assert.AreEqual("ro", crediti);
            }
        }
        [Test]
        public void TestXPathOrdering()
        {

            XSSFWorkbook wb = XSSFTestDataSamples
                    .OpenSampleWorkbook("CustomXmlMappings-inverse-order.xlsx");

            MapInfo mapInfo = null;

            foreach (POIXMLDocumentPart p in wb.GetRelations())
            {

                if (p is MapInfo)
                {
                    mapInfo = (MapInfo)p;

                    XSSFMap map = mapInfo.GetXSSFMapById(1);
                    XSSFExportToXml exporter = new XSSFExportToXml(map);

                    Assert.AreEqual(1, exporter.Compare("/CORSO/DOCENTE", "/CORSO/NOME"));
                    Assert.AreEqual(-1, exporter.Compare("/CORSO/NOME", "/CORSO/DOCENTE"));
                }
            }
        }
        [Test]
        public void TestMultiTable()
        {

            XSSFWorkbook wb = XSSFTestDataSamples
                    .OpenSampleWorkbook("CustomXMLMappings-complex-type.xlsx");

            foreach (POIXMLDocumentPart p in wb.GetRelations())
            {

                if (p is MapInfo)
                {
                    MapInfo mapInfo = (MapInfo)p;

                    XSSFMap map = mapInfo.GetXSSFMapById(2);

                    Assert.IsNotNull(map);

                    XSSFExportToXml exporter = new XSSFExportToXml(map);
                    MemoryStream os = new MemoryStream();
                    exporter.ExportToXML(os, true);
                    String xml = Encoding.UTF8.GetString(os.ToArray()); 
                    Assert.IsNotNull(xml);

                    String[] regexConditions = {
						"<MapInfo", "</MapInfo>",
						"<Schema ID=\"1\" SchemaRef=\"\" Namespace=\"\" />",
						"<Schema ID=\"4\" SchemaRef=\"\" Namespace=\"\" />",
						"DataBinding",
						"Map ID=\"1\"",
						"Map ID=\"5\"",
				};

                    foreach (String condition in regexConditions)
                    {
                        Regex pattern = new Regex(condition,RegexOptions.Compiled);
                        Assert.IsTrue(pattern.IsMatch(xml),condition);
                    }
                }
            }
        }
    }
}


