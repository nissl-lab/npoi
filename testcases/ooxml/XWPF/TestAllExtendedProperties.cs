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

namespace TestCases.XWPF
{
    using NPOI.XWPF.UserModel;
    using NPOI.OpenXmlFormats;
    using NUnit.Framework;

    /**
     * Tests if the {@link CoreProperties#getKeywords()} method. This test has been
     * submitted because even though the
     * {@link PackageProperties#getKeywordsProperty()} had been present before, the
     * {@link CoreProperties#getKeywords()} had been missing.
     * 
     * The author of this has Added {@link CoreProperties#getKeywords()} and
     * {@link CoreProperties#setKeywords(String)} and this test is supposed to test
     * them.
     * 
     * @author Antoni Mylka
     * 
     */
    [TestFixture]
    public class TestAllExtendedProperties
    {
        [Test]
        public void TestGetAllExtendedProperties()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("TestPoiXMLDocumentCorePropertiesGetKeywords.docx");
            CT_ExtendedProperties ctProps = doc.GetProperties().ExtendedProperties.GetUnderlyingProperties();
            Assert.AreEqual("Microsoft Office Word", ctProps.Application);
            Assert.AreEqual("14.0000", ctProps.AppVersion);
            Assert.AreEqual(57, ctProps.Characters);
            Assert.AreEqual(66, ctProps.CharactersWithSpaces);
            Assert.AreEqual("", ctProps.Company);
            Assert.IsNull(ctProps.DigSig);
            Assert.AreEqual(0, ctProps.DocSecurity);
            //Assert.IsNotNull(ctProps.DomNode);

            CT_VectorVariant vec = ctProps.HeadingPairs;
            Assert.AreEqual(2, vec.vector.SizeOfVariantArray());
            Assert.AreEqual("Title", vec.vector.GetVariantArray(0).lpstr);
            Assert.AreEqual(1, vec.vector.GetVariantArray(1).i4);

            Assert.IsFalse(ctProps.IsSetHiddenSlides());
            Assert.AreEqual(0, ctProps.HiddenSlides);
            Assert.IsFalse(ctProps.IsSetHLinks());
            Assert.IsNull(ctProps.HLinks);
            Assert.IsNull(ctProps.HyperlinkBase);
            Assert.IsTrue(ctProps.IsSetHyperlinksChanged());
            Assert.IsFalse(ctProps.HyperlinksChanged);
            Assert.AreEqual(1, ctProps.Lines);
            Assert.IsTrue(ctProps.IsSetLinksUpToDate());
            Assert.IsFalse(ctProps.LinksUpToDate);
            Assert.IsNull(ctProps.Manager);
            Assert.IsFalse(ctProps.IsSetMMClips());
            Assert.AreEqual(0, ctProps.MMClips);
            Assert.IsFalse(ctProps.IsSetNotes());
            Assert.AreEqual(0, ctProps.Notes);
            Assert.AreEqual(1, ctProps.Pages);
            Assert.AreEqual(1, ctProps.Paragraphs);
            Assert.IsNull(ctProps.PresentationFormat);
            Assert.IsTrue(ctProps.IsSetScaleCrop());
            Assert.IsFalse(ctProps.ScaleCrop);
            Assert.IsTrue(ctProps.IsSetSharedDoc());
            Assert.IsFalse(ctProps.SharedDoc);
            Assert.IsFalse(ctProps.IsSetSlides());
            Assert.AreEqual(0, ctProps.Slides);
            Assert.AreEqual("Normal.dotm", ctProps.Template);

            CT_VectorLpstr vec2 = ctProps.TitlesOfParts;
            Assert.AreEqual(1, vec2.vector.SizeOfLpstrArray());
            Assert.AreEqual("Example Word 2010 Document", vec2.vector.GetLpstrArray(0));

            Assert.AreEqual(3, ctProps.TotalTime);
            Assert.AreEqual(10, ctProps.Words);

            // Check the digital signature part
            // Won't be there in this file, but we
            //  need to do this check so that the
            //  appropriate parts end up in the
            //  smaller ooxml schemas file
            CT_DigSigBlob blob = ctProps.DigSig;
            Assert.IsNull(blob);

            blob = new CT_DigSigBlob();
            blob.blob = (new byte[] { 2, 6, 7, 2, 3, 4, 5, 1, 2, 3 });
        }
    }

}