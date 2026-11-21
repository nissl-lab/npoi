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
    using NUnit.Framework;using NUnit.Framework.Legacy;

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
            ClassicAssert.AreEqual("Microsoft Office Word", ctProps.Application);
            ClassicAssert.AreEqual("14.0000", ctProps.AppVersion);
            ClassicAssert.AreEqual(57, ctProps.Characters);
            ClassicAssert.AreEqual(66, ctProps.CharactersWithSpaces);
            ClassicAssert.AreEqual("", ctProps.Company);
            ClassicAssert.IsNull(ctProps.DigSig);
            ClassicAssert.AreEqual(0, ctProps.DocSecurity);
            //ClassicAssert.IsNotNull(ctProps.DomNode);

            CT_VectorVariant vec = ctProps.HeadingPairs;
            ClassicAssert.AreEqual(2, vec.vector.SizeOfVariantArray());
            ClassicAssert.AreEqual("Title", vec.vector.GetVariantArray(0).lpstr);
            ClassicAssert.AreEqual(1, vec.vector.GetVariantArray(1).i4);

            ClassicAssert.IsFalse(ctProps.IsSetHiddenSlides());
            ClassicAssert.AreEqual(0, ctProps.HiddenSlides);
            ClassicAssert.IsFalse(ctProps.IsSetHLinks());
            ClassicAssert.IsNull(ctProps.HLinks);
            ClassicAssert.IsNull(ctProps.HyperlinkBase);
            ClassicAssert.IsTrue(ctProps.IsSetHyperlinksChanged());
            ClassicAssert.IsFalse(ctProps.HyperlinksChanged);
            ClassicAssert.AreEqual(1, ctProps.Lines);
            ClassicAssert.IsTrue(ctProps.IsSetLinksUpToDate());
            ClassicAssert.IsFalse(ctProps.LinksUpToDate);
            ClassicAssert.IsNull(ctProps.Manager);
            ClassicAssert.IsFalse(ctProps.IsSetMMClips());
            ClassicAssert.AreEqual(0, ctProps.MMClips);
            ClassicAssert.IsFalse(ctProps.IsSetNotes());
            ClassicAssert.AreEqual(0, ctProps.Notes);
            ClassicAssert.AreEqual(1, ctProps.Pages);
            ClassicAssert.AreEqual(1, ctProps.Paragraphs);
            ClassicAssert.IsNull(ctProps.PresentationFormat);
            ClassicAssert.IsTrue(ctProps.IsSetScaleCrop());
            ClassicAssert.IsFalse(ctProps.ScaleCrop);
            ClassicAssert.IsTrue(ctProps.IsSetSharedDoc());
            ClassicAssert.IsFalse(ctProps.SharedDoc);
            ClassicAssert.IsFalse(ctProps.IsSetSlides());
            ClassicAssert.AreEqual(0, ctProps.Slides);
            ClassicAssert.AreEqual("Normal.dotm", ctProps.Template);

            CT_VectorLpstr vec2 = ctProps.TitlesOfParts;
            ClassicAssert.AreEqual(1, vec2.vector.SizeOfLpstrArray());
            ClassicAssert.AreEqual("Example Word 2010 Document", vec2.vector.GetLpstrArray(0));

            ClassicAssert.AreEqual(3, ctProps.TotalTime);
            ClassicAssert.AreEqual(10, ctProps.Words);

            // Check the digital signature part
            // Won't be there in this file, but we
            //  need to do this check so that the
            //  appropriate parts end up in the
            //  smaller ooxml schemas file
            CT_DigSigBlob blob = ctProps.DigSig;
            ClassicAssert.IsNull(blob);

            blob = new CT_DigSigBlob();
            blob.blob = (new byte[] { 2, 6, 7, 2, 3, 4, 5, 1, 2, 3 });
        }
    }

}