/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace TestCases.OpenXml4Net.OPC.Compliance
{

    using System;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.OpenXml4Net.Exceptions;
    using NUnit.Framework;
    using System.IO;
    using NPOI.Util;



    /**
     * Test core properties Open Packaging Convention compliance.
     * 
     * M4.1: The format designer shall specify and the format producer shall create
     * at most one core properties relationship for a package. A format consumer
     * shall consider more than one core properties relationship for a package to be
     * an error. If present, the relationship shall target the Core Properties part.
     * (POI relaxes this on reading, as Office sometimes breaks this)
     * 
     * M4.2: The format designer shall not specify and the format producer shall not
     * create Core Properties that use the Markup Compatibility namespace as defined
     * in Annex F, "Standard Namespaces and Content Types". A format consumer shall
     * consider the use of the Markup Compatibility namespace to be an error.
     * 
     * M4.3: Producers shall not create a document element that contains refinements
     * to the Dublin Core elements, except for the two specified in the schema:
     * <dcterms:created> and <dcterms:modified> Consumers shall consider a document
     * element that violates this constraint to be an error.
     * 
     * M4.4: Producers shall not create a document element that contains the
     * xml:lang attribute. Consumers shall consider a document element that violates
     * this constraint to be an error.
     * 
     * M4.5: Producers shall not create a document element that contains the
     * xsi:type attribute, except for a <dcterms:created> or <dcterms:modified>
     * element where the xsi:type attribute shall be present and shall hold the
     * value dcterms:W3CDTF, where dcterms is the namespace prefix of the Dublin
     * Core namespace. Consumers shall consider a document element that violates
     * this constraint to be an error.
     * 
     * @author Julien Chable
     */
    [TestFixture]
    public class TestOPCComplianceCoreProperties
    {

        [Test]
        public void TestCorePropertiesPart()
        {
            OPCPackage pkg;
            string path = OpenXml4NetTestDataSamples.GetSampleFileName("OPCCompliance_CoreProperties_OnlyOneCorePropertiesPart.docx");
            pkg = OPCPackage.Open(path);

            pkg.Revert();
        }
        private static String ExtractInvalidFormatMessage(String sampleNameSuffix)
        {

            Stream is1 = OpenXml4NetTestDataSamples.OpenComplianceSampleStream("OPCCompliance_CoreProperties_" + sampleNameSuffix);
            OPCPackage pkg;
            try
            {
                pkg = OPCPackage.Open(is1);
            }
            catch (InvalidFormatException e)
            {
                // no longer required for successful test
                return e.Message;
            }

            pkg.Revert();
            throw new AssertionException("expected OPC compliance exception was not thrown");
        }

        /**
         * Test M4.1 rule.
         */
        [Test]
        public void TestOnlyOneCorePropertiesPart()
        {
            // We have relaxed this check, so we can read the file anyway
            try
            {
                ExtractInvalidFormatMessage("OnlyOneCorePropertiesPartFAIL.docx");
                Assert.Fail("M4.1 should be being relaxed");
            }
            catch (AssertionException e) { }

            // We will use the first core properties, and ignore the others
            Stream is1 = OpenXml4NetTestDataSamples.OpenSampleStream("MultipleCoreProperties.docx");
            OPCPackage pkg = OPCPackage.Open(is1);

            // We can see 2 by type
            Assert.AreEqual(2, pkg.GetPartsByContentType(ContentTypes.CORE_PROPERTIES_PART).Count);
            // But only the first one by relationship
            Assert.AreEqual(1, pkg.GetPartsByRelationshipType(PackageRelationshipTypes.CORE_PROPERTIES).Count);
            // It should be core.xml not the older core1.xml
            Assert.AreEqual(
                  "/docProps/core.xml",
                  pkg.GetPartsByRelationshipType(PackageRelationshipTypes.CORE_PROPERTIES)[0].PartName.ToString()
            );
        }

        private static Uri CreateURI(String text)
        {
            return new Uri(text,UriKind.RelativeOrAbsolute);

        }

        /**
         * Test M4.1 rule.
         */
        [Test]
        public void TestOnlyOneCorePropertiesPart_AddRelationship()
        {
            Stream is1 = OpenXml4NetTestDataSamples.OpenComplianceSampleStream("OPCCompliance_CoreProperties_OnlyOneCorePropertiesPart.docx");
            OPCPackage pkg;
            pkg = OPCPackage.Open(is1);

            Uri partUri = CreateURI("/docProps/core2.xml");
            try
            {
                pkg.AddRelationship(PackagingUriHelper.CreatePartName(partUri), TargetMode.Internal,
                        PackageRelationshipTypes.CORE_PROPERTIES);
                // no longer fail on compliance error
                //fail("expected OPC compliance exception was not thrown");
            }
            catch (InvalidFormatException e)
            {
                throw;
            }
            catch (InvalidOperationException e)
            {
                // expected during successful test
                Assert.AreEqual("OPC Compliance error [M4.1]: can't add another core properties part ! Use the built-in package method instead.", e.Message);
            }
            pkg.Revert();
        }

        /**
         * Test M4.1 rule.
         */
        [Test]
        public void TestOnlyOneCorePropertiesPart_AddPart()
        {
            String sampleFileName = "OPCCompliance_CoreProperties_OnlyOneCorePropertiesPart.docx";
            OPCPackage pkg = null;
            pkg = OPCPackage.Open(POIDataSamples.GetOpenXml4NetInstance().GetFile(sampleFileName));


            Uri partUri = CreateURI("/docProps/core2.xml");
            try
            {
                pkg.CreatePart(PackagingUriHelper.CreatePartName(partUri),
                        ContentTypes.CORE_PROPERTIES_PART);
                // no longer fail on compliance error
                //fail("expected OPC compliance exception was not thrown");
            }
            catch (InvalidFormatException e)
            {
                throw;
            }
            catch (InvalidOperationException e)
            {
                // expected during successful test
                Assert.AreEqual("OPC Compliance error [M4.1]: you try to add more than one core properties relationship in the package !", e.Message);
            }
            pkg.Revert();
        }

        /**
         * Test M4.2 rule.
         */
        [Test]
        public void TestDoNotUseCompatibilityMarkup()
        {
            String msg = ExtractInvalidFormatMessage("DoNotUseCompatibilityMarkupFAIL.docx");
            Assert.AreEqual("OPC Compliance error [M4.2]: A format consumer shall consider the use of the Markup Compatibility namespace to be an error.", msg);
        }

        /**
         * Test M4.3 rule.
         */
        [Test]
        public void TestDCTermsNamespaceLimitedUse()
        {
            String msg = ExtractInvalidFormatMessage("DCTermsNamespaceLimitedUseFAIL.docx");
            Assert.AreEqual("OPC Compliance error [M4.3]: Producers shall not create a document element that contains refinements to the Dublin Core elements, except for the two specified in the schema: <dcterms:created> and <dcterms:modified> Consumers shall consider a document element that violates this constraint to be an error.", msg);
        }

        /**
         * Test M4.4 rule.
         */
        [Test]
        public void TestUnauthorizedXMLLangAttribute()
        {
            String msg = ExtractInvalidFormatMessage("UnauthorizedXMLLangAttributeFAIL.docx");
            Assert.AreEqual("OPC Compliance error [M4.4]: Producers shall not create a document element that contains the xml:lang attribute. Consumers shall consider a document element that violates this constraint to be an error.", msg);
        }

        /**
         * Test M4.5 rule.
         */
        [Test]
        public void TestLimitedXSITypeAttribute_NotPresent()
        {
            String msg = ExtractInvalidFormatMessage("LimitedXSITypeAttribute_NotPresentFAIL.docx");
            Assert.AreEqual("The element 'created' must have the 'xsi:type' attribute present !", msg);
        }

        /**
         * Test M4.5 rule.
         */
        [Test]
        public void TestLimitedXSITypeAttribute_PresentWithUnauthorizedValue()
        {
            String msg = ExtractInvalidFormatMessage("LimitedXSITypeAttribute_PresentWithUnauthorizedValueFAIL.docx");
            Assert.AreEqual("The element 'modified' must have the 'xsi:type' attribute with the value 'dcterms:W3CDTF' !", msg);
        }

        /**
     * Document with no core properties - testing at the OPC level,
     *  saving into a new stream
     */
        [Test]
        public void TestNoCoreProperties_saveNew()
        {
            String sampleFileName = "OPCCompliance_NoCoreProperties.xlsx";
            OPCPackage pkg = OPCPackage.Open(POIDataSamples.GetOpenXml4NetInstance().GetFileInfo(sampleFileName).FullName);

            // Verify it has empty properties
            Assert.AreEqual(0, pkg.GetPartsByContentType(ContentTypes.CORE_PROPERTIES_PART).Count);
            Assert.IsNotNull(pkg.GetPackageProperties());
            Assert.IsNull(pkg.GetPackageProperties().GetLanguageProperty());
            //Assert.IsNull(pkg.GetPackageProperties().GetLanguageProperty().Value);

            // Save and re-load
            MemoryStream baos = new MemoryStream();
            pkg.Save(baos);
            MemoryStream bais = new MemoryStream(baos.ToArray());

            pkg = OPCPackage.Open(bais);

            // An Empty Properties part has been Added in the save/load
            Assert.AreEqual(1, pkg.GetPartsByContentType(ContentTypes.CORE_PROPERTIES_PART).Count);
            Assert.IsNotNull(pkg.GetPackageProperties());
            Assert.IsNull(pkg.GetPackageProperties().GetLanguageProperty());
            //Assert.IsNull(pkg.GetPackageProperties().GetLanguageProperty().Value);

            // Open a new copy of it
            pkg = OPCPackage.Open(POIDataSamples.GetOpenXml4NetInstance().GetFileInfo(sampleFileName).FullName);

            // Save and re-load, without having touched the properties yet
            baos = new MemoryStream();
            pkg.Save(baos);
            bais = new MemoryStream(baos.ToArray());
            pkg = OPCPackage.Open(bais);

            // Check that this too Added empty properties without error
            Assert.AreEqual(1, pkg.GetPartsByContentType(ContentTypes.CORE_PROPERTIES_PART).Count);
            Assert.IsNotNull(pkg.GetPackageProperties());
            Assert.IsNull(pkg.GetPackageProperties().GetLanguageProperty());
            //Assert.IsNull(pkg.GetPackageProperties().GetLanguageProperty().Value);

        }

        /**
         * Document with no core properties - testing at the OPC level,
         *  from a temp-file, saving in-place
         */
        [Test]
        public void TestNoCoreProperties_saveInPlace()
        {
            String sampleFileName = "OPCCompliance_NoCoreProperties.xlsx";

            // Copy this into a temp file, so we can play with it
            FileInfo tmp = TempFile.CreateTempFile("poi-test", ".opc");
            FileStream out1 = new FileStream(tmp.FullName, FileMode.Create, FileAccess.ReadWrite);
            IOUtils.Copy(
                    POIDataSamples.GetOpenXml4NetInstance().OpenResourceAsStream(sampleFileName),
                    out1);
            out1.Close();

            // Open it from that temp file
            OPCPackage pkg = OPCPackage.Open(tmp);

            // Empty properties
            Assert.AreEqual(0, pkg.GetPartsByContentType(ContentTypes.CORE_PROPERTIES_PART).Count);
            Assert.IsNotNull(pkg.GetPackageProperties());
            Assert.IsNull(pkg.GetPackageProperties().GetLanguageProperty());
            //Assert.IsNull(pkg.GetPackageProperties().GetLanguageProperty().Value);

            // Save and close
            pkg.Close();


            // Re-open and check
            pkg = OPCPackage.Open(tmp);

            // An Empty Properties part has been Added in the save/load
            Assert.AreEqual(1, pkg.GetPartsByContentType(ContentTypes.CORE_PROPERTIES_PART).Count);
            Assert.IsNotNull(pkg.GetPackageProperties());
            Assert.IsNull(pkg.GetPackageProperties().GetLanguageProperty());
            //Assert.IsNull(pkg.GetPackageProperties().GetLanguageProperty().Value);

            // Finish and tidy
            pkg.Revert();
            tmp.Delete();
        }

    }
}