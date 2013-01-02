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

namespace NPOI.XWPF.UserModel
{

    using NUnit.Framework;

    using NPOI.XWPF;
    using NPOI.OpenXmlFormats.Wordprocessing;


    [TestFixture]
    public class TestXWPFFootnotes
    {

        [Test]
        public void TestAddFootnotesToDocument()
        {
            XWPFDocument docOut = new XWPFDocument();

            int noteId = 1;

            XWPFFootnotes footnotes = docOut.CreateFootnotes();
            CT_FtnEdn ctNote = new CT_FtnEdn();
            ctNote.id = (noteId.ToString());
            ctNote.type = (ST_FtnEdn.normal);
            footnotes.AddFootnote(ctNote);

            XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(docOut);

            XWPFFootnote note = docIn.GetFootnoteByID(noteId);
            Assert.AreEqual(note.GetCTFtnEdn().type, ST_FtnEdn.normal);
        }
    }

}
