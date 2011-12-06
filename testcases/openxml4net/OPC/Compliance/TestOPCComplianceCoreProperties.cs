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
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.IO;



    /**
     * Test core properties Open Packaging Convention compliance.
     * 
     * M4.1: The format designer shall specify and the format producer shall create
     * at most one core properties relationship for a package. A format consumer
     * shall consider more than one core properties relationship for a package to be
     * an error. If present, the relationship shall target the Core Properties part.
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
    [TestClass]
    public class TestOPCComplianceCoreProperties
    {

        [TestMethod]
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
                // expected during successful test
                return e.Message;
            }

            pkg.Revert();
            // Normally must thrown an InvalidFormatException exception.
            throw new AssertFailedException("expected OPC compliance exception was not thrown");
        }

        /**
         * Test M4.1 rule.
         */
        [TestMethod]
        public void TestOnlyOneCorePropertiesPart()
        {
            String msg = ExtractInvalidFormatMessage("OnlyOneCorePropertiesPartFAIL.docx");
            Assert.AreEqual("OPC Compliance error [M4.1]: there is more than one core properties relationship in the package !", msg);
        }

        private static Uri CreateURI(String text)
        {
            return new Uri(text,UriKind.RelativeOrAbsolute);

        }

        /**
         * Test M4.1 rule.
         */
        [TestMethod]
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
                Assert.Fail("expected OPC compliance exception was not thrown");
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
        [TestMethod]
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
                Assert.Fail("expected OPC compliance exception was not thrown");
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
        [TestMethod]
        public void TestDoNotUseCompatibilityMarkup()
        {
            String msg = ExtractInvalidFormatMessage("DoNotUseCompatibilityMarkupFAIL.docx");
            Assert.AreEqual("OPC Compliance error [M4.2]: A format consumer shall consider the use of the Markup Compatibility namespace to be an error.", msg);
        }

        /**
         * Test M4.3 rule.
         */
        [TestMethod]
        public void TestDCTermsNamespaceLimitedUse()
        {
            String msg = ExtractInvalidFormatMessage("DCTermsNamespaceLimitedUseFAIL.docx");
            Assert.AreEqual("OPC Compliance error [M4.3]: Producers shall not create a document element that contains refinements to the Dublin Core elements, except for the two specified in the schema: <dcterms:created> and <dcterms:modified> Consumers shall consider a document element that violates this constraint to be an error.", msg);
        }

        /**
         * Test M4.4 rule.
         */
        [TestMethod]
        public void TestUnauthorizedXMLLangAttribute()
        {
            String msg = ExtractInvalidFormatMessage("UnauthorizedXMLLangAttributeFAIL.docx");
            Assert.AreEqual("OPC Compliance error [M4.4]: Producers shall not create a document element that contains the xml:lang attribute. Consumers shall consider a document element that violates this constraint to be an error.", msg);
        }

        /**
         * Test M4.5 rule.
         */
        [TestMethod]
        public void TestLimitedXSITypeAttribute_NotPresent()
        {
            String msg = ExtractInvalidFormatMessage("LimitedXSITypeAttribute_NotPresentFAIL.docx");
            Assert.AreEqual("The element 'created' must have the 'xsi:type' attribute present !", msg);
        }

        /**
         * Test M4.5 rule.
         */
        [TestMethod]
        public void TestLimitedXSITypeAttribute_PresentWithUnauthorizedValue()
        {
            String msg = ExtractInvalidFormatMessage("LimitedXSITypeAttribute_PresentWithUnauthorizedValueFAIL.docx");
            Assert.AreEqual("The element 'modified' must have the 'xsi:type' attribute with the value 'dcterms:W3CDTF' !", msg);
        }
    }
}