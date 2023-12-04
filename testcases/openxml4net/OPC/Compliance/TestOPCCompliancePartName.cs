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
    using NUnit.Framework;
    using NPOI.OpenXml4Net.Exceptions;
    /**
     * Test part name Open Packaging Convention compliance.
     *
     * (Open Packaging Convention 8.1.1 Part names) :
     *
     * The part name grammar is defined as follows:
     *
     * part_name = 1*( "/" segment )
     *
     * segment = 1*( pchar )
     *
     * pchar is defined in RFC 3986.
     *
     * The part name grammar implies the following constraints. The package
     * implementer shall neither create any part that violates these constraints nor
     * retrieve any data from a package as a part if the purported part name
     * violates these constraints.
     *
     * A part name shall not be empty. [M1.1]
     *
     * A part name shall not have empty segments. [M1.3]
     *
     * A part name shall start with a forward slash ("/") character. [M1.4]
     *
     * A part name shall not have a forward slash as the last character. [M1.5]
     *
     * A segment shall not hold any characters other than pchar characters. [M1.6]
     *
     * Part segments have the following additional constraints. The package
     * implementer shall neither create any part with a part name comprised of a
     * segment that violates these constraints nor retrieve any data from a package
     * as a part if the purported part name contains a segment that violates these
     * constraints.
     *
     * A segment shall not contain percent-encoded forward slash ("/"), or backward
     * slash ("\") characters. [M1.7]
     *
     * A segment shall not contain percent-encoded unreserved characters. [M1.8]
     *
     * A segment shall not end with a dot (".") character. [M1.9]
     *
     * A segment shall include at least one non-dot character. [M1.10]
     *
     * A package implementer shall neither create nor recognize a part with a part
     * name derived from another part name by appending segments to it. [M1.11]
     *
     * Part name equivalence is determined by comparing part names as
     * case-insensitive ASCII strings. [M1.12]
     *
     * @author Julien Chable
     */
    [TestFixture]
    public class TestOPCCompliancePartName
    {

        /**
         * Test some common invalid names.
         *
         * A segment shall not contain percent-encoded unreserved characters. [M1.8]
         */
        [Test]
        public void TestInvalidPartNames()
        {
            String[] invalidNames = { "/", "/xml./doc.xml", "[Content_Types].xml", "//xml/." };
            foreach (String s in invalidNames)
            {
                Uri uri = null;
                try
                {
                    uri = new Uri(s, UriKind.Relative);
                }
                catch
                {
                    Assert.IsTrue(s.Equals("[Content_Types].xml"));
                    continue;
                }
                Assert.IsFalse(
                        PackagingUriHelper.IsValidPartName(uri), "This part name SHOULD NOT be valid: " + s);
            }
        }

        /**
         * Test some common valid names.
         */
        [Test]
        public void TestValidPartNames()
        {
            String[] validNames = { "/xml/item1.xml", "/document.xml",
                "/a/%D1%86.xml" };
            foreach (String s in validNames)
                Assert.IsTrue(
                        PackagingUriHelper.IsValidPartName(new Uri(s, UriKind.RelativeOrAbsolute)),
                        "This part name SHOULD be valid: " + s);
        }

        /**
         * A part name shall not be empty. [M1.1]
         */
        [Test]
        public void TestEmptyPartNameFailure()
        {
            try
            {
                PackagingUriHelper.CreatePartName(new Uri("", UriKind.Relative));
                Assert.Fail("A part name shall not be empty. [M1.1]");
            }
            catch (InvalidFormatException)
            {
                // Normal behaviour
            }
        }

        /**
         * A part name shall not have empty segments. [M1.3]
         *
         * A segment shall not end with a dot ('.') character. [M1.9]
         *
         * A segment shall include at least one non-dot character. [M1.10]
         */
        [Test]
        public void TestPartNameWithInvalidSegmentsFailure()
        {
            String[] invalidNames = { "//document.xml", "//word/document.xml",
                "/word//document.rels", "/word//rels//document.rels",
                "/xml./doc.xml", "/document.", "/./document.xml",
                "/word/./doc.rels", "/%2F/document.xml" };
            foreach (String s in invalidNames)
                Assert.IsFalse(
                        PackagingUriHelper.IsValidPartName(new Uri(s, UriKind.RelativeOrAbsolute)),
                        "A part name shall not have empty segments. [M1.3]");
        }

        /**
         * A segment shall not hold any characters other than ipchar (RFC 3987) characters.
         * [M1.6].
         */
        [Test]
        public void TestPartNameWithNonPCharCharacters()
        {
            String[] validNames = { "/doc&.xml" };
            try
            {
                foreach (String s in validNames)
                    Assert.IsTrue(
                            PackagingUriHelper
                                    .IsValidPartName(new Uri(s, UriKind.RelativeOrAbsolute)),
                                    "A segment shall not contain non pchar characters [M1.6] : "
                                    + s);
            }
            catch (UriFormatException e)
            {
                Assert.Fail(e.Message);
            }
        }

        /**
         * A segment shall not contain percent-encoded unreserved characters [M1.8].
         */
        [Test]
        public void TestPartNameWithUnreservedEncodedCharactersFailure()
        {
            String[] invalidNames = { "/a/docum%65nt.xml" };
            try
            {
                foreach (String s in invalidNames)
                    Assert.IsFalse(
                             PackagingUriHelper
                                    .IsValidPartName(new Uri(s, UriKind.RelativeOrAbsolute)),
                                    "A segment shall not contain percent-encoded unreserved characters [M1.8] : "
                                    + s);
            }
            catch (UriFormatException e)
            {
                Assert.Fail(e.Message);
            }
        }

        /**
         * A part name shall start with a forward slash ('/') character. [M1.4]
         */
        [Test]
        public void TestPartNameStartsWithAForwardSlashFailure()
        {
            try
            {
                PackagingUriHelper.CreatePartName(new Uri("document.xml", UriKind.RelativeOrAbsolute));
                Assert.Fail("A part name shall start with a forward slash ('/') character. [M1.4]");
            }
            catch (InvalidFormatException)
            {
                // Normal behaviour
            }
        }

        /**
         * A part name shall not have a forward slash as the last character. [M1.5]
         */
        [Test]
        public void TestPartNameEndsWithAForwardSlashFailure()
        {
            try
            {
                PackagingUriHelper.CreatePartName(new Uri("/document.xml/", UriKind.Relative));
                Assert.Fail("A part name shall not have a forward slash as the last character. [M1.5]");
            }
            catch (InvalidFormatException)
            {
                // Normal behaviour
            }
        }

        /**
         * Part name equivalence is determined by comparing part names as
         * case-insensitive ASCII strings. [M1.12]
         */
        [Test]
        public void TestPartNameComparaison()
        {
            String[] partName1 = { "/word/document.xml", "/docProps/core.xml", "/rels/.rels" };
            String[] partName2 = { "/WORD/DocUment.XML", "/docProps/core.xml", "/rels/.rels" };
            for (int i = 0; i < partName1.Length || i < partName2.Length; ++i)
            {
                PackagePartName p1 = PackagingUriHelper.CreatePartName(partName1[i]);
                PackagePartName p2 = PackagingUriHelper.CreatePartName(partName2[i]);
                Assert.IsTrue(p1.Equals(p2));
                Assert.IsTrue(p1.CompareTo(p2) == 0);
                Assert.IsTrue(p1.GetHashCode() == p2.GetHashCode());
            }
        }

        /**
         * Part name equivalence is determined by comparing part names as
         * case-insensitive ASCII strings. [M1.12].
         *
         * All the comparisons MUST FAIL !
         */
        [Test]
        public void TestPartNameComparaisonFailure()
        {
            String[] partName1 = { "/word/document.xml", "/docProps/core.xml", "/rels/.rels" };
            String[] partName2 = { "/WORD/DocUment.XML2", "/docProp/core.xml", "/rels/rels" };
            for (int i = 0; i < partName1.Length || i < partName2.Length; ++i)
            {
                PackagePartName p1 = PackagingUriHelper.CreatePartName(partName1[i]);
                PackagePartName p2 = PackagingUriHelper.CreatePartName(partName2[i]);
                Assert.IsFalse(p1.Equals(p2));
                Assert.IsFalse(p1.CompareTo(p2) == 0);
                Assert.IsFalse(p1.GetHashCode() == p2.GetHashCode());
            }
        }
    }
}