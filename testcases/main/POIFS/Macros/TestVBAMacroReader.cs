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

namespace NPOI.POIFS.macros
{
    using System;

using static NPOI.POITESTCASE;
using static NPOI.POITESTCASE;
using static NPOI.POITESTCASE;
using static org.junit.Assert;
using static org.junit.Assert;









using org.apache.poi;
using NPOI.POIFS.FileSystem;
using NPOI.UTIL;
using NPOI.UTIL;
using org.junit;
using org.junit;

    [TestClass]
    public class TestVBAMacroReader
{
    private static Dictionary<POIDataSamples, String> expectedMacroContents;

    protected static String ReadVBA(POIDataSamples poiDataSamples) {
        File macro = poiDataSamples.GetFile("SimpleMacro.vba");
        byte[] bytes;
        try {
            FileInputStream stream = new FileInputStream(macro);
            try {
                bytes = IOUtils.ToByteArray(stream);
            } finally {
                stream.Close();
            }
        } catch (IOException e) {
            throw new Exception(e);
        }

        String testMacroContents = new String(bytes, StringUtil.UTF8);
        
        if (! testMacroContents.StartsWith("Sub ")) {
            throw new ArgumentException("Not a macro");
        }

        return testMacroContents.Substring(testMacroContents.IndexOf("()")+3);
    }

    static {
        Dictionary<POIDataSamples, String> _expectedMacroContents = new Dictionary<POIDataSamples, String>();
        POIDataSamples[] dataSamples = {
                POIDataSamples.SpreadSheetInstance,
                POIDataSamples.SlideShowInstance,
                POIDataSamples.DocumentInstance,
                POIDataSamples.DiagramInstance
        };
        foreach (POIDataSamples sample in dataSamples) {
            _expectedMacroContents.Put(sample, ReadVBA(sample));
        }
        expectedMacroContents = Collections.UnmodifiableMap(_expectedMacroContents);
    }
    
    //////////////////////////////// From Stream /////////////////////////////
    [Test]
    public void HSSFfromStream()  {
        fromStream(POIDataSamples.SpreadSheetInstance, "SimpleMacro.xls");
    }
    [Test]
    public void XSSFfromStream()  {
        fromStream(POIDataSamples.SpreadSheetInstance, "SimpleMacro.xlsm");
    }
    @Ignore("bug 59302: Found 0 macros")
    [Test]
    public void HSLFfromStream()  {
        fromStream(POIDataSamples.SlideShowInstance, "SimpleMacro.ppt");
    }
    [Test]
    public void XSLFfromStream()  {
        fromStream(POIDataSamples.SlideShowInstance, "SimpleMacro.pptm");
    }
    [Test]
    public void HWPFfromStream()  {
        fromStream(POIDataSamples.DocumentInstance, "SimpleMacro.doc");
    }
    [Test]
    public void XWPFfromStream()  {
        fromStream(POIDataSamples.DocumentInstance, "SimpleMacro.docm");
    }
    @Ignore("Found 0 macros")
    [Test]
    public void HDGFfromStream()  {
        fromStream(POIDataSamples.DiagramInstance, "SimpleMacro.vsd");
    }
    [Test]
    public void XDGFfromStream()  {
        fromStream(POIDataSamples.DiagramInstance, "SimpleMacro.vsdm");
    }

    //////////////////////////////// From File /////////////////////////////
    [Test]
    public void HSSFfromFile()  {
        fromFile(POIDataSamples.SpreadSheetInstance, "SimpleMacro.xls");
    }
    [Test]
    public void XSSFfromFile()  {
        fromFile(POIDataSamples.SpreadSheetInstance, "SimpleMacro.xlsm");
    }
    @Ignore("bug 59302: Found 0 macros")
    [Test]
    public void HSLFfromFile()  {
        fromFile(POIDataSamples.SlideShowInstance, "SimpleMacro.ppt");
    }
    [Test]
    public void XSLFfromFile()  {
        fromFile(POIDataSamples.SlideShowInstance, "SimpleMacro.pptm");
    }
    [Test]
    public void HWPFfromFile()  {
        fromFile(POIDataSamples.DocumentInstance, "SimpleMacro.doc");
    }
    [Test]
    public void XWPFfromFile()  {
        fromFile(POIDataSamples.DocumentInstance, "SimpleMacro.docm");
    }
    @Ignore("Found 0 macros")
    [Test]
    public void HDGFfromFile()  {
        fromFile(POIDataSamples.DiagramInstance, "SimpleMacro.vsd");
    }
    [Test]
    public void XDGFfromFile()  {
        fromFile(POIDataSamples.DiagramInstance, "SimpleMacro.vsdm");
    }

    //////////////////////////////// From NPOIFS /////////////////////////////
    [Test]
    public void HSSFfromNPOIFS()  {
        fromNPOIFS(POIDataSamples.SpreadSheetInstance, "SimpleMacro.xls");
    }
    @Ignore("bug 59302: Found 0 macros")
    [Test]
    public void HSLFfromNPOIFS()  {
        fromNPOIFS(POIDataSamples.SlideShowInstance, "SimpleMacro.ppt");
    }
    [Test]
    public void HWPFfromNPOIFS()  {
        fromNPOIFS(POIDataSamples.DocumentInstance, "SimpleMacro.doc");
    }
    @Ignore("Found 0 macros")
    [Test]
    public void HDGFfromNPOIFS()  {
        fromNPOIFS(POIDataSamples.DiagramInstance, "SimpleMacro.vsd");
    }

    protected void fromFile(POIDataSamples dataSamples, String filename) throws IOException {
        File f = dataSamples.GetFile(filename);
        VBAMacroReader r = new VBAMacroReader(f);
        try {
            assertMacroContents(dataSamples, r);
        } finally {
            r.Close();
        }
    }

    protected void fromStream(POIDataSamples dataSamples, String filename) throws IOException {
        InputStream fis = dataSamples.OpenResourceAsStream(filename);
        try {
            VBAMacroReader r = new VBAMacroReader(fis);
            try {
                assertMacroContents(dataSamples, r);
            } finally {
                r.Close();
            }
        } finally {
            fis.Close();
        }
    }

    protected void fromNPOIFS(POIDataSamples dataSamples, String filename) throws IOException {
        File f = dataSamples.GetFile(filename);
        NPOIFSFileSystem fs = new NPOIFSFileSystem(f);
        try {
            VBAMacroReader r = new VBAMacroReader(fs);
            try {
                assertMacroContents(dataSamples, r);
            } finally {
                r.Close();
            }
        } finally {
            fs.Close();
        }
    }
    
    protected void assertMacroContents(POIDataSamples samples, VBAMacroReader r) throws IOException {
        Assert.IsNotNull(r);
        Dictionary<String,String> contents = r.ReadMacros();
        Assert.IsNotNull(contents);
        Assert.IsFalse("Found 0 macros", contents.IsEmpty());
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
        assertContains(contents, "Module1");
        String content = contents.Get("Module1");
        Assert.IsNotNull(content);
        assertContains(content, "Attribute VB_Name = \"Module1\"");
        //assertContains(content, "Attribute TestMacro.VB_Description = \"This is a test macro\"");

        // And the macro itself
        String testMacroNoSub = expectedMacroContents.Get(samples);
        assertContains(content, testMacroNoSub);
    }
    
    @Ignore
    [Test]
    public void bug59830() throws IOException {
        // This file is intentionally omitted from the test-data directory
        // unless we can extract the vbaProject.bin from this Word 97-2003 file
        // so that it's less likely to be opened and executed on a Windows computer.
        // The file is attached to bug 59830.
        // The Macro Virus only affects Windows computers, as it Makes a
        // subprocess call to powershell.exe with an encoded payload
        // The document Contains macros that execute on workbook open if macros
        // are enabled
        File doc = POIDataSamples.DocumentInstance.GetFile("macro_virus.doc.do_not_open");
        VBAMacroReader Reader = new VBAMacroReader(doc);
        Dictionary<String, String> macros = Reader.ReadMacros();
        Assert.IsNotNull(macros);
        Reader.Close();
    }
    
    // This test is written as expected-to-fail and should be rewritten
    // as expected-to-pass when the bug is fixed.
    [Test]
    public void bug59858() throws IOException {
        try {
            fromFile(POIDataSamples.SpreadSheetInstance, "59858.xls");
            testPassesNow(59858);
        } catch (IOException e) {
            if (e.Message.Matches("Module offset for '.+' was never Read.")) {
                //e.PrintStackTrace();
                // NPE when Reading module.offset in VBAMacroReader.ReadMacros (approx line 258)
                skipTest(e);
            } else {
                // something unexpected failed
                throw e;
            }
        }
    }
    
    // This test is written as expected-to-fail and should be rewritten
    // as expected-to-pass when the bug is fixed.
    [Test]
    public void bug60158() throws IOException {
        try {
            fromFile(POIDataSamples.DocumentInstance, "60158.docm");
            testPassesNow(60158);
        } catch (ArrayIndexOutOfBoundsException e) {
            skipTest(e);
        }
    }
}

