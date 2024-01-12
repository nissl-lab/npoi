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
using NPOI.Util;
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
using System.Xml.Serialization;

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

            foreach(POIXMLDocumentPart p in wb.GetRelations())
            {

                if(!(p is MapInfo))
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

            foreach(POIXMLDocumentPart p in wb.GetRelations())
            {

                if(!(p is MapInfo))
                {
                    continue;
                }
                mapInfo = (MapInfo) p;

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

            foreach(POIXMLDocumentPart p in wb.GetRelations())
            {

                if(p is MapInfo)
                {
                    mapInfo = (MapInfo) p;

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

            foreach(POIXMLDocumentPart p in wb.GetRelations())
            {

                if(p is MapInfo)
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

                    foreach(String condition in regexConditions)
                    {
                        Regex pattern = new Regex(condition,RegexOptions.Compiled);
                        Assert.IsTrue(pattern.IsMatch(xml), condition);
                    }
                }
            }
        }
        [Test]
        public void Test55850ComplexXmlExport()
        {

            XSSFWorkbook wb = XSSFTestDataSamples
                .OpenSampleWorkbook("55850.xlsx");

            bool found = false;
            foreach(POIXMLDocumentPart p in wb.GetRelations())
            {

                if(!(p is MapInfo))
                {
                    continue;
                }
                MapInfo mapInfo = (MapInfo) p;

                XSSFMap map = mapInfo.GetXSSFMapById(2);

                Assert.IsNotNull(map, "XSSFMap is null");

                XSSFExportToXml exporter = new XSSFExportToXml(map);
                ByteArrayOutputStream os = new ByteArrayOutputStream();
                exporter.ExportToXML(os, true);
                String xmlData = os.ToString("UTF-8");

                Assert.IsNotNull(xmlData);
                Assert.AreNotEqual("", xmlData);

                String a = xmlData.Split(new string[] {"<A>" },StringSplitOptions.None)[1]
                    .Split(new string[] {"</A>" },StringSplitOptions.None)[0].Trim();
                String b = a.Split(new string[] {"<B>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</B>" }, StringSplitOptions.None)[0].Trim();
                String c = b.Split(new string[] {"<C>}" },StringSplitOptions.None)[0]
                    .Split(new string[] {"</C>" },StringSplitOptions.None)[0].Trim();
                String d = c.Split(new string[] {"<D>" },StringSplitOptions.None)[1]
                    .Split(new string[] {"</Dd>" },StringSplitOptions.None)[0].Trim();
                String e = d.Split(new string[] { "<E>" },StringSplitOptions.None)[1]
                    .Split(new string[] { "</EA>" },StringSplitOptions.None)[0].Trim();
                String euro = e.Split(new string[] {"<EUR>" },StringSplitOptions.None)[1]
                    .Split(new string[] {"</EUR>" },StringSplitOptions.None)[0].Trim();
                String chf = e.Split(new string[] {"<CHF>" },StringSplitOptions.None)[1]
                    .Split(new string[] {"</CHF>" },StringSplitOptions.None)[0].Trim();

                Assert.AreEqual("15", euro);
                Assert.AreEqual("19", chf);

                parseXML(xmlData);

                found = true;
            }
            Assert.True(found);
        }
        [Test]
        public void testFormulaCells_Bugzilla_55927()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("55927.xlsx");

            bool found = false;
            foreach(POIXMLDocumentPart p in wb.GetRelations())
            {
                if(!(p is MapInfo))
                {
                    continue;
                }
                MapInfo mapInfo = (MapInfo) p;

                XSSFMap map = mapInfo.GetXSSFMapById(1);

                Assert.IsNotNull(map, "XSSFMap is null");

                XSSFExportToXml exporter = new XSSFExportToXml(map);
                ByteArrayOutputStream os = new ByteArrayOutputStream();
                exporter.ExportToXML(os, true);
                String xmlData = os.ToString("UTF-8");

                Assert.IsNotNull(xmlData);
                Assert.AreNotEqual("", xmlData);

                Assert.AreEqual("2012-01-13", xmlData.Split(new string[] { "<DATE>" }, StringSplitOptions.None)[1]
                    .Split(new string[] { "</DATE>" }, StringSplitOptions.None)[0].Trim());
                Assert.AreEqual("2012-02-16", xmlData.Split(new string[] { "<FORMULA_DATE>" }, StringSplitOptions.None)[1]
                    .Split(new string[]{"</FORMULA_DATE>"}, StringSplitOptions.None)[0].Trim());

                parseXML(xmlData);

                found = true;
            }
            Assert.IsTrue(found);
        }
        [Test]
        public void testFormulaCells_Bugzilla_55926()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("55926.xlsx");

            bool found = false;
            foreach(POIXMLDocumentPart p in wb.GetRelations())
            {
                if(!(p is MapInfo))
                {
                    continue;
                }
                MapInfo mapInfo = (MapInfo) p;

                XSSFMap map = mapInfo.GetXSSFMapById(1);

                Assert.IsNotNull(map, "XSSFMap is null");

                XSSFExportToXml exporter = new XSSFExportToXml(map);
                ByteArrayOutputStream os = new ByteArrayOutputStream();
                exporter.ExportToXML(os, true);
                String xmlData = os.ToString("UTF-8");

                Assert.IsNotNull(xmlData);
                Assert.AreNotEqual("", xmlData);

                String a = xmlData.Split(new string[] {"<A>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</A>"}, StringSplitOptions.None)[0].Trim();
                String doubleValue = a.Split(new string[] {"<DOUBLE>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</DOUBLE>" }, StringSplitOptions.None)[0].Trim();
                String stringValue = a.Split(new string[] {"<STRING>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</STRING>"}, StringSplitOptions.None)[0].Trim();

                Assert.AreEqual("Hello World", stringValue);
                Assert.AreEqual("5.0999999999999996", doubleValue);

                parseXML(xmlData);

                found = true;
            }
            Assert.IsTrue(found);
        }
        [Test]
        public void testXmlExportIgnoresEmptyCells_Bugzilla_55924()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("55924.xlsx");

            bool found = false;
            foreach(POIXMLDocumentPart p in wb.GetRelations())
            {
                if(!(p is MapInfo))
                {
                    continue;
                }
                MapInfo mapInfo = (MapInfo) p;

                XSSFMap map = mapInfo.GetXSSFMapById(1);

                Assert.IsNotNull(map, "XSSFMap is null");

                XSSFExportToXml exporter = new XSSFExportToXml(map);
                ByteArrayOutputStream os = new ByteArrayOutputStream();
                exporter.ExportToXML(os, true);
                String xmlData = os.ToString("UTF-8");

                Assert.IsNotNull(xmlData);
                Assert.AreNotEqual("", xmlData);

                String a =  xmlData.Split(new string[] {"<A>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</A>" }, StringSplitOptions.None)[0].Trim();
                String euro = a.Split(new string[] {"<EUR>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</EUR>" }, StringSplitOptions.None)[0].Trim();
                Assert.AreEqual("1", euro);

                parseXML(xmlData);

                found = true;
            }
            Assert.IsTrue(found);
        }
        [Test]
        public void testXmlExportCompare_Bug_55923()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("55923.xlsx");

            bool found = false;
            foreach(POIXMLDocumentPart p in wb.GetRelations())
            {
                if(!(p is MapInfo))
                {
                    continue;
                }
                MapInfo mapInfo = (MapInfo) p;

                XSSFMap map = mapInfo.GetXSSFMapById(4);

                Assert.IsNotNull(map, "XSSFMap is null");

                XSSFExportToXml exporter = new XSSFExportToXml(map);
                Assert.AreEqual(0, exporter.Compare("", ""));
                Assert.AreEqual(0, exporter.Compare("/", "/"));
                Assert.AreEqual(0, exporter.Compare("//", "//"));
                Assert.AreEqual(0, exporter.Compare("/a/", "/b/"));
                Assert.AreEqual(-1, exporter.Compare("/ns1:Entry/ns1:A/ns1:B/ns1:C/ns1:E/ns1:EUR",
                                                "/ns1:Entry/ns1:A/ns1:B/ns1:C/ns1:E/ns1:CHF"));

                found = true;
            }
            Assert.IsTrue(found);
        }
        [Test]
        public void testXmlExportSchemaWithXSAllTag_Bugzilla_56169()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56169.xlsx");

            bool found = false;
            foreach(XSSFMap map in wb.GetCustomXMLMappings())
            {
                XSSFExportToXml exporter = new XSSFExportToXml(map);
                ByteArrayOutputStream os = new ByteArrayOutputStream();
                exporter.ExportToXML(os, true);
                String xmlData = os.ToString("UTF-8");

                Assert.IsNotNull(xmlData);
                Assert.AreNotEqual("", xmlData);

                String a = xmlData.Split(new string[] {"<A>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</A>" }, StringSplitOptions.None)[0].Trim();
                String a_b = a.Split(new string[] {"<B>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</B>" }, StringSplitOptions.None)[0].Trim();
                String a_b_c = a_b.Split(new string[] {"<C>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</C>" }, StringSplitOptions.None)[0].Trim();
                String a_b_c_e = a_b_c.Split(new string[] {"<E>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</EA>" }, StringSplitOptions.None)[0].Trim();
                String a_b_c_e_euro = a_b_c_e.Split(new string[] {"<EUR>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</EUR>" }, StringSplitOptions.None)[0].Trim();
                String a_b_c_e_chf = a_b_c_e.Split(new string[] {"<CHF>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</CHF>" }, StringSplitOptions.None)[0].Trim();

                Assert.AreEqual("1", a_b_c_e_euro);
                Assert.AreEqual("2", a_b_c_e_chf);

                String a_b_d = a_b.Split(new string[] {"<D>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</D>" }, StringSplitOptions.None)[0].Trim();
                String a_b_d_e = a_b_d.Split(new string[] {"<E>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</E>" }, StringSplitOptions.None)[0].Trim();

                String a_b_d_e_euro = a_b_d_e.Split(new string[] {"<EUR>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</EUR>" }, StringSplitOptions.None)[0].Trim();
                String a_b_d_e_chf = a_b_d_e.Split(new string[] {"<CHF>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</CHF>" }, StringSplitOptions.None)[0].Trim();

                Assert.AreEqual("3", a_b_d_e_euro);
                Assert.AreEqual("4", a_b_d_e_chf);

                found = true;
            }
            Assert.IsTrue(found);
        }
        [Test]
        public void testXmlExportSchemaOrderingBug_Bugzilla_55923()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("55923.xlsx");

            bool found = false;
            foreach(POIXMLDocumentPart p in wb.GetRelations())
            {
                if(!(p is MapInfo))
                {
                    continue;
                }
                MapInfo mapInfo = (MapInfo) p;

                XSSFMap map = mapInfo.GetXSSFMapById(4);

                Assert.IsNotNull(map, "XSSFMap is null");

                XSSFExportToXml exporter = new XSSFExportToXml(map);
                ByteArrayOutputStream os = new ByteArrayOutputStream();
                exporter.ExportToXML(os, true);
                String xmlData = os.ToString("UTF-8");

                Assert.IsNotNull(xmlData);
                Assert.AreNotEqual("", xmlData);

                String a = xmlData.Split(new string[] {"<A>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</A>" }, StringSplitOptions.None)[0].Trim();
                String a_b = a.Split(new string[] {"<B>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</B>" }, StringSplitOptions.None)[0].Trim();
                String a_b_c = a_b.Split(new string[] {"<C>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</C>" }, StringSplitOptions.None)[0].Trim();
                String a_b_c_e = a_b_c.Split(new string[] {"<E>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</EA>" }, StringSplitOptions.None)[0].Trim();
                String a_b_c_e_euro = a_b_c_e.Split(new string[] {"<EUR>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</EUR>" }, StringSplitOptions.None)[0].Trim();
                String a_b_c_e_chf = a_b_c_e.Split(new string[] {"<CHF>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</CHF>" }, StringSplitOptions.None)[0].Trim();

                Assert.AreEqual("1", a_b_c_e_euro);
                Assert.AreEqual("2", a_b_c_e_chf);

                String a_b_d = a_b.Split(new string[] {"<D>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</D>" }, StringSplitOptions.None)[0].Trim();
                String a_b_d_e = a_b_d.Split(new string[] {"<E>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</E>" }, StringSplitOptions.None)[0].Trim();

                String a_b_d_e_euro = a_b_d_e.Split(new string[] {"<EUR>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</EUR>" }, StringSplitOptions.None)[0].Trim();
                String a_b_d_e_chf = a_b_d_e.Split(new string[] {"<CHF>" }, StringSplitOptions.None)[1]
                    .Split(new string[] {"</CHF>" }, StringSplitOptions.None)[0].Trim();

                Assert.AreEqual("3", a_b_d_e_euro);
                Assert.AreEqual("4", a_b_d_e_chf);

                found = true;
            }
            Assert.IsTrue(found);
        }
        private void parseXML(String xmlData) {
            string _byteOrderMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());
            if(xmlData.StartsWith(_byteOrderMarkUtf8))
            {
                xmlData = xmlData.Remove(0, _byteOrderMarkUtf8.Length);
            }
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlData);
        }
    }
}


