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
namespace TestCases.XWPF.UserModel
{
    using ICSharpCode.SharpZipLib.Zip;
    using NPOI;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using NPOI.Util;
    using NPOI.XWPF;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;
    using System;
    using System.Xml;
    using TestCases;

    [TestFixture]
    public class TestXWPFBugs
    {
        [Test]
        public void Bug55802()
        {
            String blabla =
                "Bir, iki, \u00fc\u00e7, d\u00f6rt, be\u015f,\n" +
                "\nalt\u0131, yedi, sekiz, dokuz, on.\n" +
                "\nK\u0131rm\u0131z\u0131 don,\n" +
                "\ngel bizim bah\u00e7eye kon,\n" +
                "\nsar\u0131 limon";
            XWPFDocument doc = new XWPFDocument();
            XWPFRun run = doc.CreateParagraph().CreateRun();

            foreach (String str in blabla.Split("\n".ToCharArray()))
            {
                run.SetText(str);
                run.AddBreak();
            }

            run.FontFamily = (/*setter*/"Times New Roman");
            run.FontSize = (/*setter*/20);
            Assert.AreEqual(run.FontFamily, "Times New Roman");
            Assert.AreEqual(run.GetFontFamily(FontCharRange.CS), "Times New Roman");
            Assert.AreEqual(run.GetFontFamily(FontCharRange.EastAsia), "Times New Roman");
            Assert.AreEqual(run.GetFontFamily(FontCharRange.HAnsi), "Times New Roman");
            run.SetFontFamily("Arial", FontCharRange.HAnsi);
            Assert.AreEqual(run.GetFontFamily(FontCharRange.HAnsi), "Arial");

            doc.Close();
        }

        [Test]
        public void Bug57312_NullPointException()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("57312.docx");
            Assert.IsNotNull(doc);

            foreach (IBodyElement bodyElement in doc.BodyElements)
            {
                BodyElementType elementType = bodyElement.ElementType;

                if (elementType == BodyElementType.PARAGRAPH)
                {
                    XWPFParagraph paragraph = (XWPFParagraph)bodyElement;

                    foreach (IRunElement iRunElem in paragraph.IRuns)
                    {

                        if (iRunElem is XWPFRun)
                        {
                            XWPFRun RunElement = (XWPFRun)iRunElem;

                            UnderlinePatterns underline = RunElement.Underline;
                            Assert.IsNotNull(underline);

                            //System.out.Println("Found: " + underline + ": " + RunElement.GetText(0));
                        }
                    }
                }
            }
        }

        [Test]
        public void Bug57495_getTableArrayInDoc()
        {
            XWPFDocument doc = new XWPFDocument();
            //let's create a few tables for the test
            for (int i = 0; i < 3; i++)
            {
                doc.CreateTable(2, 2);
            }
            XWPFTable table = doc.GetTableArray(0);
            Assert.IsNotNull(table);
            //let's check also that returns the correct table
            XWPFTable same = doc.Tables[0];
            Assert.AreEqual(table, same);
        }

        [Test]
        public void Bug57495_getParagraphArrayInTableCell()
        {
            XWPFDocument doc = new XWPFDocument();
            //let's create a table for the test
            XWPFTable table = doc.CreateTable(2, 2);
            Assert.IsNotNull(table);
            XWPFParagraph p = table.GetRow(0).GetCell(0).GetParagraphArray(0);
            Assert.IsNotNull(p);
            //let's check also that returns the correct paragraph
            XWPFParagraph same = table.GetRow(0).GetCell(0).Paragraphs[0];
            Assert.AreEqual(p, same);
        }

        [Test]
        public void Bug57495_convertPixelsToEMUs()
        {
            int pixels = 100;
            int expectedEMU = 952500;
            int result = Units.PixelToEMU(pixels);
            Assert.AreEqual(expectedEMU, result);
        }


        [Test]
        public void Test56392()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("56392.docx");
            Assert.IsNotNull(doc);
        }
        /**
         * Removing a run needs to remove it from both Runs and IRuns
         */
        [Test]
        public void Test57829()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            Assert.IsNotNull(doc);
            Assert.AreEqual(3, doc.Paragraphs.Count);
            foreach (XWPFParagraph paragraph in doc.Paragraphs)
            {
                paragraph.RemoveRun(0);
                Assert.IsNotNull(paragraph.Text);
            }
        }

        /**
         * Removing a run needs to take into account position of run if paragraph contains hyperlink runs
         */
        [Test]
        public void Test58618()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("58618.docx");
            XWPFParagraph para = (XWPFParagraph)doc.BodyElements[0];
            Assert.IsNotNull(para);
            Assert.AreEqual("Some text  some hyper links link link and some text.....", para.Text);
            XWPFRun run = para.InsertNewRun(para.Runs.Count);
            run.SetText("New Text");
            Assert.AreEqual("Some text  some hyper links link link and some text.....New Text", para.Text);
            para.RemoveRun(para.Runs.Count - 2);
            Assert.AreEqual("Some text  some hyper links link linkNew Text", para.Text);
        }
        [Test]
        public void Bug59058()
        {
            String[] files = { "bug57031.docx", "bug59058.docx" };
            foreach (String f in files)
            {
                ZipFile zf = new ZipFile(POIDataSamples.GetDocumentInstance().GetFile(f));
                ZipEntry entry = zf.GetEntry("word/document.xml");
                XmlDocument xml = POIXMLDocumentPart.ConvertStreamToXml(zf.GetInputStream(entry));
                DocumentDocument document = DocumentDocument.Parse(xml, POIXMLDocumentPart.NamespaceManager);
                Assert.IsNotNull(document);
                zf.Close();
            }
        }

        [Test]
        public void Test59378()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("59378.docx");
            ByteArrayOutputStream out1 = new ByteArrayOutputStream();
            doc.Write(out1);
            out1.Close();
            XWPFDocument doc2 = new XWPFDocument(new ByteArrayInputStream(out1.ToByteArray()));
            doc2.Close();
            XWPFDocument docBack = XWPFTestDataSamples.WriteOutAndReadBack(doc);
            docBack.Close();
        }
    }

}