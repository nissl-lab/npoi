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
    using System;

    using NPOI.XWPF;
    using NUnit.Framework;
    using System.IO;
    using NPOI.XWPF.UserModel;

    [TestFixture]
    public class TestChangeTracking
    {

        [Test]
        public void Detection()
        {

            XWPFDocument documentWithoutChangeTracking = XWPFTestDataSamples.OpenSampleDocument("bug56075-changeTracking_off.docx");
            Assert.IsFalse(documentWithoutChangeTracking.IsTrackRevisions);

            XWPFDocument documentWithChangeTracking = XWPFTestDataSamples.OpenSampleDocument("bug56075-changeTracking_on.docx");
            Assert.IsTrue(documentWithChangeTracking.IsTrackRevisions);

        }

        [Test]
        public void ActivateChangeTracking()
        {
            XWPFDocument document = XWPFTestDataSamples.OpenSampleDocument("bug56075-changeTracking_off.docx");
            Assert.IsFalse(document.IsTrackRevisions);

            document.IsTrackRevisions = (/*setter*/true);

            Assert.IsTrue(document.IsTrackRevisions);
        }

        [Test]
        public void Integration()
        {
            XWPFDocument doc = new XWPFDocument();

            XWPFParagraph p1 = doc.CreateParagraph();

            XWPFRun r1 = p1.CreateRun();
            r1.SetText("Lorem ipsum dolor sit amet.");
            doc.IsTrackRevisions = (true);

            MemoryStream out1 = new MemoryStream();
            doc.Write(out1);

            MemoryStream inputStream = new MemoryStream(out1.ToArray());
            XWPFDocument document = new XWPFDocument(inputStream);
            inputStream.Close();

            Assert.IsTrue(document.IsTrackRevisions);
        }

    }
}

