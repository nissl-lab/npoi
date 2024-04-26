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

namespace TestCases.POIFS.Macros
{
    using NPOI.POIFS.FileSystem;
    using NPOI.POIFS.Macros;
    using NPOI.Util;
    using NUnit.Framework;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using System.Text.RegularExpressions;
    using TestCases;

    [TestFixture]
    [Platform("Win", Reason = "Expected to run on Windows platform")]
    public class TestVBAMacroReader
    {
        private static IReadOnlyDictionary<POIDataSamples, String> expectedMacroContents;

        protected static String ReadVBA(POIDataSamples poiDataSamples)
        {
            FileInfo macro = poiDataSamples.GetFileInfo("SimpleMacro.vba");
            byte[] bytes;
            try
            {
                FileStream stream = new FileStream(macro.FullName, FileMode.Open, FileAccess.ReadWrite);
                try
                {
                    bytes = IOUtils.ToByteArray(stream);
                }
                finally
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                //throw new Exception(e);
                throw;
            }

            String testMacroContents = Encoding.UTF8.GetString(bytes);

            if (!testMacroContents.StartsWith("Sub "))
            {
                throw new ArgumentException("Not a macro");
            }

            return testMacroContents.Substring(testMacroContents.IndexOf("()") + 3);
        }

        static TestVBAMacroReader()
        {
            Dictionary<POIDataSamples, String> _expectedMacroContents = new Dictionary<POIDataSamples, String>();
            POIDataSamples[] dataSamples = {
                POIDataSamples.GetSpreadSheetInstance(),
                POIDataSamples.GetSlideShowInstance(),
                POIDataSamples.GetDocumentInstance(),
                POIDataSamples.GetDiagramInstance()
        };
            foreach (POIDataSamples sample in dataSamples)
            {
                _expectedMacroContents.Add(sample, ReadVBA(sample));
            }
            expectedMacroContents = _expectedMacroContents;
        }

        //////////////////////////////// From Stream /////////////////////////////
        [Test]
        public void HSSFfromStream()
        {
            FromStream(POIDataSamples.GetSpreadSheetInstance(), "SimpleMacro.xls");
        }

        [Test]
        public void XSSFfromStream()
        {
            FromStream(POIDataSamples.GetSpreadSheetInstance(), "SimpleMacro.xlsm");
        }

        [Ignore("bug 59302: Found 0 macros")]
        [Test]
        public void HSLFfromStream()
        {
            FromStream(POIDataSamples.GetSlideShowInstance(), "SimpleMacro.ppt");
        }

        [Ignore("not support XSLF")]
        [Test]
        public void XSLFfromStream()
        {
            FromStream(POIDataSamples.GetSlideShowInstance(), "SimpleMacro.pptm");
        }

        [Test]
        public void HWPFfromStream()
        {
            FromStream(POIDataSamples.GetDocumentInstance(), "SimpleMacro.doc");
        }

        [Test]
        public void XWPFfromStream()
        {
            FromStream(POIDataSamples.GetDocumentInstance(), "SimpleMacro.docm");
        }

        [Ignore("Found 0 macros")]
        [Test]
        public void HDGFfromStream()
        {
            FromStream(POIDataSamples.GetDiagramInstance(), "SimpleMacro.vsd");
        }

        [Ignore("not support XDGF")]
        [Test]
        public void XDGFfromStream()
        {
            FromStream(POIDataSamples.GetDiagramInstance(), "SimpleMacro.vsdm");
        }

        //////////////////////////////// From File /////////////////////////////
        [Test]
        public void HSSFfromFile()
        {
            FromFile(POIDataSamples.GetSpreadSheetInstance(), "SimpleMacro.xls");
        }

        [Test]
        public void XSSFfromFile()
        {
            FromFile(POIDataSamples.GetSpreadSheetInstance(), "SimpleMacro.xlsm");
        }

        [Ignore("bug 59302: Found 0 macros")]
        [Test]
        public void HSLFfromFile()
        {
            FromFile(POIDataSamples.GetSlideShowInstance(), "SimpleMacro.ppt");
        }

        [Ignore("not support XSLF")]
        [Test]
        public void XSLFfromFile()
        {
            FromFile(POIDataSamples.GetSlideShowInstance(), "SimpleMacro.pptm");
        }

        [Test]
        public void HWPFfromFile()
        {
            FromFile(POIDataSamples.GetDocumentInstance(), "SimpleMacro.doc");
        }

        [Test]
        public void XWPFfromFile()
        {
            FromFile(POIDataSamples.GetDocumentInstance(), "SimpleMacro.docm");
        }

        [Ignore("Found 0 macros")]
        [Test]
        public void HDGFfromFile()
        {
            FromFile(POIDataSamples.GetDiagramInstance(), "SimpleMacro.vsd");
        }

        [Ignore("not support XDGF")]
        [Test]
        public void XDGFfromFile()
        {
            FromFile(POIDataSamples.GetDiagramInstance(), "SimpleMacro.vsdm");
        }

        //////////////////////////////// From NPOIFS /////////////////////////////
        [Test]
        public void HSSFfromNPOIFS()
        {
            FromNPOIFS(POIDataSamples.GetSpreadSheetInstance(), "SimpleMacro.xls");
        }

        [Ignore("bug 59302: Found 0 macros")]
        [Test]
        public void HSLFfromNPOIFS()
        {
            FromNPOIFS(POIDataSamples.GetSlideShowInstance(), "SimpleMacro.ppt");
        }

        [Test]
        public void HWPFfromNPOIFS()
        {
            FromNPOIFS(POIDataSamples.GetDocumentInstance(), "SimpleMacro.doc");
        }

        [Ignore("Found 0 macros")]
        [Test]
        public void HDGFfromNPOIFS()
        {
            FromNPOIFS(POIDataSamples.GetDiagramInstance(), "SimpleMacro.vsd");
        }

        protected void FromFile(POIDataSamples dataSamples, String filename)
        {
            FileInfo f = dataSamples.GetFileInfo(filename);
            VBAMacroReader r = new VBAMacroReader(f);
            try
            {
                AssertMacroContents(dataSamples, r);
            }
            finally
            {
                r.Close();
            }
        }

        protected void FromStream(POIDataSamples dataSamples, String filename)
        {
            InputStream fis = new FileInputStream(dataSamples.OpenResourceAsStream(filename));
            try
            {
                VBAMacroReader r = new VBAMacroReader(fis);
                try
                {
                    AssertMacroContents(dataSamples, r);
                }
                finally
                {
                    r.Close();
                }
            }
            finally
            {
                fis.Close();
            }
        }

        protected void FromNPOIFS(POIDataSamples dataSamples, String filename)
        {
            FileInfo f = dataSamples.GetFileInfo(filename);
            NPOIFSFileSystem fs = new NPOIFSFileSystem(f);
            try
            {
                VBAMacroReader r = new VBAMacroReader(fs);
                try
                {
                    AssertMacroContents(dataSamples, r);
                }
                finally
                {
                    r.Close();
                }
            }
            finally
            {
                fs.Close();
            }
        }

        protected void AssertMacroContents(POIDataSamples samples, VBAMacroReader r)
        {
            Assert.IsNotNull(r);
            Dictionary<String, String> contents = r.ReadMacros();
            Assert.IsNotNull(contents);
            Assert.IsFalse(contents.Count == 0, "Found 0 macros");
            /*
            Assert.AreEqual(5, contents.Size());

            // Check the ones without scripts
            String[] noScripts = new String[] { "ThisWorkbook",
                    "Sheet1", "Sheet2", "Sheet3" };
            foreach (String entry in noScripts) {
                Assert.IsTrue(entry, contents.ContainsKey(entry));

                String content = contents.Get(entry);
                assertContains(content, "Attribute VB_Exposed = True");
                assertContains(content, "Attribute VB_Customizable = True");
                assertContains(content, "Attribute VB_TemplateDerived = False");
                assertContains(content, "Attribute VB_GlobalNameSpace = False");
                assertContains(content, "Attribute VB_Exposed = True");
            }
            */

            // Check the script one
            POITestCase.AssertContains(contents, "Module1");
            String content = contents["Module1"];
            Assert.IsNotNull(content);
            POITestCase.AssertContains(content, "Attribute VB_Name = \"Module1\"");
            //assertContains(content, "Attribute TestMacro.VB_Description = \"This is a test macro\"");

            // And the macro itself
            String testMacroNoSub = expectedMacroContents[samples];
            POITestCase.AssertContains(content, testMacroNoSub);
        }

        [Test]
        public void Bug59830()
        {
            //test file is "609751.xls" in govdocs1
            FileInfo f = POIDataSamples.GetSpreadSheetInstance().GetFileInfo("59830.xls");
            VBAMacroReader r = new VBAMacroReader(f);
            Dictionary<string, string> macros = r.ReadMacros();
            Assert.IsNotNull(macros["Module20"]);
            StringAssert.Contains("here start of superscripting", macros["Module20"]);
            r.Close();
        }

        // This test is written as expected-to-fail and should be rewritten
        // as expected-to-pass when the bug is fixed.
        [Test]
        public void Bug59858()
        {
            FileInfo f = POIDataSamples.GetSpreadSheetInstance().GetFileInfo("59858.xls");
            VBAMacroReader r = new VBAMacroReader(f);
            Dictionary<string, string> macros = r.ReadMacros();
            Assert.IsNotNull(macros["Sheet4"]);
            StringAssert.Contains("intentional constituent", macros["Sheet4"]);
            r.Close();
        }

        // This test is written as expected-to-fail and should be rewritten
        // as expected-to-pass when the bug is fixed.
        [Test]
        public void Bug60158()
        {
            FileInfo f = POIDataSamples.GetDocumentInstance().GetFileInfo("60158.docm");
            VBAMacroReader r = new VBAMacroReader(f);
            Dictionary<string, string> macros = r.ReadMacros();
            Assert.IsNotNull(macros["NewMacros"]);
            StringAssert.Contains("' dirty", macros["NewMacros"]);
            r.Close();
        }

        [Test]
        public void Bug60273()
        {
            //test file derives from govdocs1 147240.xls
            FileInfo f = POIDataSamples.GetSpreadSheetInstance().GetFileInfo("60273.xls");
            VBAMacroReader r = new VBAMacroReader(f);
            Dictionary<string, string> macros = r.ReadMacros();
            Assert.IsNotNull(macros["Module1"]);
            StringAssert.Contains("9/8/2004", macros["Module1"]);
            r.Close();
        }
    }
}