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
    using NPOI.OpenXmlFormats.Wordprocessing;
    using NPOI.XWPF;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;
    using System;


    /**
     * @author Paolo Mottadelli
     */
    [TestFixture]
    public class TestXWPFHeadings
    {
        private static String HEADING1 = "Heading1";

        [Test]
        public void TestSetParagraphStyle()
        {
            //new clean instance of paragraph
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("heading123.docx");
            XWPFParagraph p = doc.CreateParagraph();
            XWPFRun run = p.CreateRun();
            run.SetText("Heading 1");

            CT_SdtBlock block = doc.Document.body.AddNewSdt();

            Assert.IsNull(p.Style);
            p.Style = HEADING1;
            Assert.AreEqual(HEADING1, p.GetCTP().pPr.pStyle.val);

            doc.CreateTOC();
            /*
            // TODO - finish this test
            if (false) {
                CTStyles styles = doc.Style;
                CTStyle style = styles.AddNewStyle();
                style.Type=(STStyleType.PARAGRAPH);
                style.StyleId=("Heading1");
            }

            if (false) {
                File file = TempFile.CreateTempFile("testHeaders", ".docx");
                OutputStream out1 = new FileOutputStream(file);
                doc.Write(out1);
                out1.Close();
            }
            */
        }
    }

}