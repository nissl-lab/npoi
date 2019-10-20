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

namespace NPOI.Openxml4j.opc
{
    using System;
    using System.IO;
    using System.Xml;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.SS.UserModel;
    using NPOI.XSSF;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;
    using TestCases.HSSF;
    using TestCases.OpenXml4Net;

    [TestFixture]
    public class TestZipPackage
    {

        [Test]
        public void TestBug56479()
        {
            Stream is1 = OpenXml4NetTestDataSamples.OpenSampleStream("dcterms_bug_56479.zip");
            OPCPackage p = OPCPackage.Open(is1);

            // Check we found the contents of it
            bool foundCoreProps = false, foundDocument = false, foundTheme1 = false;
            foreach (PackagePart part in p.GetParts())
            {
                if (part.PartName.ToString().Equals("/docProps/core.xml"))
                {
                    Assert.AreEqual(ContentTypes.CORE_PROPERTIES_PART, part.ContentType);
                    foundCoreProps = true;
                }
                if (part.PartName.ToString().Equals("/word/document.xml"))
                {
                    Assert.AreEqual(XWPFRelation.DOCUMENT.ContentType, part.ContentType);
                    foundDocument = true;
                }
                if (part.PartName.ToString().Equals("/word/theme/theme1.xml"))
                {
                    Assert.AreEqual(XWPFRelation.THEME.ContentType, part.ContentType);
                    foundTheme1 = true;
                }
            }
            Assert.IsTrue(foundCoreProps, "Core not found in " + p.GetParts());
            Assert.IsFalse(foundDocument, "Document should not be found in " + p.GetParts());
            Assert.IsFalse(foundTheme1, "Theme1 should not found in " + p.GetParts());
        }

        [Test]
        public void TestZipEntityExpansionTerminates()
        {
            try
            {
                IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("poc-xmlbomb.xlsx");
                wb.Close();
                Assert.Fail("Should catch exception due to entity expansion limitations");
            }
            catch (POIXMLException e)
            {
                assertEntityLimitReached(e);
            }
        }
        private void assertEntityLimitReached(Exception e)
        {
            //ByteArrayOutputStream str = new ByteArrayOutputStream();
            //PrintWriter writer = new PrintWriter(new OutputStreamWriter(str, "UTF-8"));
            try
            {
                //e.printStackTrace(writer);
            }
            finally
            {
                //writer.close();
            }
            //String string1 = new String(str.toByteArray(), "UTF-8");
            String string1 = e.Message;
            Assert.IsTrue(string1.Contains("Exceeded Entity dereference bytes limit"), "Had: " + string1);
        }
        [Test]
        [Ignore("not implement class ExtractorFactory")]
        public void TestZipEntityExpansionExceedsMemory()
        {
            try
            {
                IWorkbook wb = WorkbookFactory.Create(XSSFTestDataSamples.OpenSamplePackage("poc-xmlbomb.xlsx"));
                wb.Close();
                Assert.Fail("Should catch exception due to entity expansion limitations");
            }
            catch (POIXMLException e)
            {
                assertEntityLimitReached(e);
            }
            try
            {
                //POITextExtractor extractor = ExtractorFactory.CreateExtractor(HSSFTestDataSamples.GetSampleFile("poc-xmlbomb.xlsx"));
                //try
                //{
                //    Assert.IsNotNull(extractor);

                //    try
                //    {
                //        string tmp = extractor.Text;
                //    }
                //    catch (InvalidOperationException e)
                //    {
                //        // expected due to shared strings expansion
                //    }
                //}
                //finally
                //{
                //    extractor.Close();
                //}
            }
            catch (POIXMLException e)
            {
                assertEntityLimitReached(e);
            }
        }
        [Test]
        [Ignore("not implement class ExtractorFactory")]
        public void TestZipEntityExpansionSharedStringTable()
        {
            IWorkbook wb = WorkbookFactory.Create(XSSFTestDataSamples.OpenSamplePackage("poc-shared-strings.xlsx"));
            wb.Close();

            //POITextExtractor extractor = ExtractorFactory.CreateExtractor(HSSFTestDataSamples.GetSampleFile("poc-shared-strings.xlsx"));
            //try
            //{
            //    Assert.IsNotNull(extractor);
            //    try
            //    {
            //        string tmp = extractor.Text;
            //    }
            //    catch (InvalidOperationException e)
            //    {
            //        // expected due to shared strings expansion
            //    }
            //}
            //finally
            //{
            //    extractor.Close();
            //}
        }
        [Test]
        [Ignore("not implement class ExtractorFactory")]
        public void TestZipEntityExpansionSharedStringTableEvents()
        {
            //bool before = ExtractorFactory.ThreadPrefersEventExtractors;
            //ExtractorFactory.setThreadPrefersEventExtractors(true);
            try
            {
                //POITextExtractor extractor = ExtractorFactory.CreateExtractor(HSSFTestDataSamples.GetSampleFile("poc-shared-strings.xlsx"));
                //try
                //{
                //    Assert.IsNotNull(extractor);

                //    try
                //    {
                //        string tmp = extractor.Text;
                //    }
                //    catch (InvalidOperationException e)
                //    {
                //        // expected due to shared strings expansion
                //    }
                //}
                //finally
                //{
                //    extractor.Close();
                //}
            }
            catch (XmlException e)
            {
                assertEntityLimitReached(e);
            }
            finally
            {
                //ExtractorFactory.setThreadPrefersEventExtractors(before);
            }
        }

    }

}