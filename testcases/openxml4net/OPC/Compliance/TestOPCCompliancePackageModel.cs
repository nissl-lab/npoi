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

using NPOI.OpenXml4Net.OPC;
using System;
using NPOI.OpenXml4Net.Exceptions;
using NUnit.Framework;
namespace TestCases.OpenXml4Net.OPC.Compliance
{
    /**
     * Test Open Packaging Convention package model compliance.
     *
     * M1.11 : A package implementer shall neither create nor recognize a part with
     * a part name derived from another part name by appending segments to it.
     *
     * @author Julien Chable
     */
    [TestFixture]
    public class TestOPCCompliancePackageModel
    {

        /**
         * A package implementer shall neither create nor recognize a part with a
         * part name derived from another part name by appending segments to it.
         * [M1.11]
         */
        [Test]
        public void TestPartNameDerivationAdditionAssert_Failure()
        {
            OPCPackage pkg = OPCPackage.Create("TODELETEIFEXIST.docx");
            try
            {
                PackagePartName name = PackagingUriHelper
                        .CreatePartName("/word/document.xml");
                PackagePartName nameDerived = PackagingUriHelper
                        .CreatePartName("/word/document.xml/image1.gif");
                pkg.CreatePart(name, ContentTypes.XML);
                pkg.CreatePart(nameDerived, ContentTypes.EXTENSION_GIF);
            }
            catch (InvalidOperationException)
            {
                pkg.Revert();
                return;
            }
            Assert.Fail("A package implementer shall neither create nor recognize a part with a"
                    + " part name derived from another part name by appending segments to it."
                    + " [M1.11]");
        }

        ///**
        // * A package implementer shall neither create nor recognize a part with a
        // * part name derived from another part name by appending segments to it.
        // * [M1.11]
        // */
        [Test]
        public void TestPartNameDerivationReadingAssert_Failure()
        {
            String filename = "OPCCompliance_DerivedPartNameFAIL.docx";
            try
            {
                OPCPackage.Open(POIDataSamples.GetOpenXML4JInstance().OpenResourceAsStream(filename));
            }
            catch (InvalidFormatException)
            {
                return;
            }
            Assert.Fail("A package implementer shall neither create nor recognize a part with a"
                    + " part name derived from another part name by appending segments to it."
                    + " [M1.11]");
        }

        /**
         * Rule M1.12 : Packages shall not contain equivalent part names and package
         * implementers shall neither create nor recognize packages with equivalent
         * part names.
         */
        [Test]
        public void TestAddPackageAlreadyAddFailure()
        {
            OPCPackage pkg = OPCPackage.Create("DELETEIFEXISTS.docx");
            PackagePartName name1 = null;
            PackagePartName name2 = null;
            try
            {
                name1 = PackagingUriHelper.CreatePartName("/word/document.xml");
                name2 = PackagingUriHelper.CreatePartName("/word/document.xml");
            }
            catch (InvalidFormatException e)
            {
                throw new Exception(e.Message);
            }
            pkg.CreatePart(name1, ContentTypes.XML);
            try
            {
                pkg.CreatePart(name2, ContentTypes.XML);
            }
            catch (PartAlreadyExistsException)
            {
                return;
            }
            Assert.Fail("Packages shall not contain equivalent part names and package implementers shall neither create nor recognize packages with equivalent part names. [M1.12]");
        }

        /**
         * Rule M1.12 : Packages shall not contain equivalent part names and package
         * implementers shall neither create nor recognize packages with equivalent
         * part names.
         */
        [Test]
        public void TestAddPackageAlreadyAddAssert_Failure2()
        {
            OPCPackage pkg = OPCPackage.Create("DELETEIFEXISTS.docx");
            PackagePartName partName = null;
            try
            {
                partName = PackagingUriHelper.CreatePartName("/word/document.xml");
            }
            catch (InvalidFormatException e)
            {
                throw new Exception(e.Message);
            }
            pkg.CreatePart(partName, ContentTypes.XML);
            try
            {
                pkg.CreatePart(partName, ContentTypes.XML);
            }
            catch (InvalidOperationException)
            {
                return;
            }
            Assert.Fail("Packages shall not contain equivalent part names and package implementers shall neither create nor recognize packages with equivalent part names. [M1.12]");
        }

        /**
         * Try to add a relationship to a relationship part.
         *
         * Check rule M1.25: The Relationships part shall not have relationships to
         * any other part. Package implementers shall enforce this requirement upon
         * the attempt to create such a relationship and shall treat any such
         * relationship as invalid.
         */
        [Test]
        public void TestAddRelationshipRelationshipsPartAssert_Failure()
        {
            OPCPackage pkg = OPCPackage.Create("DELETEIFEXISTS.docx");
            PackagePartName name1 = null;
            try
            {
                name1 = PackagingUriHelper
                        .CreatePartName("/test/_rels/document.xml.rels");
            }
            catch (InvalidFormatException)
            {
                Assert.Fail("This exception should never happen !");
            }

            try
            {
                pkg.AddRelationship(name1, TargetMode.Internal,
                        PackageRelationshipTypes.CORE_DOCUMENT);
            }
            catch (InvalidOperationException)
            {
                return;
            }
            Assert.Fail("Assert.Fail test -> M1.25: The Relationships part shall not have relationships to any other part");
        }
    }
}
