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

using System;
using NUnit.Framework;
using NPOI.OpenXml4Net.OPC.Internal;
using NPOI.OpenXml4Net.Exceptions;
using System.IO;
using NPOI.OpenXml4Net.OPC;
using TestCases.OpenXml4Net;
using NPOI.XWPF.UserModel;
namespace TestCase.OPC
{

    /**
     * Tests for content type (ContentType class).
     *
     * @author Julien Chable
     */
    [TestFixture]
    public class TestContentType
    {

        /**
         * Check rule M1.13: Package implementers shall only create and only
         * recognize parts with a content type; format designers shall specify a
         * content type for each part included in the format. Content types for
         * namespace parts shall fit the defInition and syntax for media types as
         * specified in RFC 2616, \u00A73.7.
         */
        [Test]
        public void TestContentTypeValidation()
        {
            String[] contentTypesToTest = new String[] { "text/xml",
                "application/pgp-key", "application/vnd.hp-PCLXL",
                "application/vnd.lotus-1-2-3" };
            for (int i = 0; i < contentTypesToTest.Length; ++i)
            {
                new ContentType(contentTypesToTest[i]);
            }
        }

        /**
         * Check rule M1.13 : Package implementers shall only create and only
         * recognize parts with a content type; format designers shall specify a
         * content type for each part included in the format. Content types for
         * namespace parts shall fit the defInition and syntax for media types as
         * specified in RFC 2616, \u00A3.7.
         *
         * Check rule M1.14: Content types shall not use linear white space either
         * between the type and subtype or between an attribute and its value.
         * Content types also shall not have leading or trailing white spaces.
         * Package implementers shall create only such content types and shall
         * require such content types when retrieving a part from a namespace; format
         * designers shall specify only such content types for inclusion in the
         * format.
         */
        [Test]
        public void TestContentTypeValidationFailure()
        {
            String[] contentTypesToTest = new String[] { "text/xml/app", "",
                "test", "text(xml/xml", "text)xml/xml", "text<xml/xml",
                "text>/xml", "text@/xml", "text,/xml", "text;/xml",
                "text:/xml", "text\\/xml", "t/ext/xml", "t\"ext/xml",
                "text[/xml", "text]/xml", "text?/xml", "tex=t/xml",
                "te{xt/xml", "tex}t/xml", "te xt/xml",
                "text" + (char) 9 + "/xml", "text xml", " text/xml " };
            for (int i = 0; i < contentTypesToTest.Length; ++i)
            {
                try
                {
                    new ContentType(contentTypesToTest[i]);
                }
                catch (InvalidFormatException e)
                {
                    continue;
                }
                Assert.Fail("Must have fail for content type: '" + contentTypesToTest[i]
                        + "' !");
            }
        }
        /**
        * Parameters are allowed, provides that they meet the
        *  criteria of rule [01.2]
        * Invalid parameters are verified as incorrect in 
        *  {@link #testContentTypeParameterFailure()}
        */
        [Test]
        public void TestContentTypeParam()
        {
            String[] contentTypesToTest = new String[] { "mail/toto;titi=tata",
               "text/xml;a=b;c=d", "text/xml;key1=param1;key2=param2",
               "application/pgp-key;version=\"2\"", 
               "application/x-resqml+xml;version=2.0;type=obj_global2dCrs"
      };
            foreach (String contentType in contentTypesToTest)
            {
                new ContentType(contentType);
            }
        }
        /**
         * Check rule [O1.2]: Format designers might restrict the usage of
         * parameters for content types.
         */
        [Test]
        public void TestContentTypeParameterFailure()
        {
            String[] contentTypesToTest = new String[] { 
                "mail/toto;\"titi=tata\"", // quotes not allowed like that
                "mail/toto;titi = tata", // spaces not allowed
                "text/\u0080" // characters above ASCII are not allowed
        };
            for (int i = 0; i < contentTypesToTest.Length; ++i)
            {
                try
                {
                    new ContentType(contentTypesToTest[i]);
                }
                catch (InvalidFormatException e)
                {
                    continue;
                }
                Assert.Fail("Must have fail for content type: '" + contentTypesToTest[i]
                        + "' !");
            }
        }

        /**
         * Check rule M1.15: The namespace implementer shall require a content type
         * that does not include comments and the format designer shall specify such
         * a content type.
         */
        [Test]
        public void TestContentTypeCommentFailure()
        {
            String[] contentTypesToTest = new String[] { "text/xml(comment)" };
            for (int i = 0; i < contentTypesToTest.Length; ++i)
            {
                try
                {
                    new ContentType(contentTypesToTest[i]);
                }
                catch (InvalidFormatException e)
                {
                    continue;
                }
                Assert.Fail("Must have fail for content type: '" + contentTypesToTest[i]
                        + "' !");
            }
        }

        /**
         * OOXML content types don't need entities, but we shouldn't
         * barf if we Get one from a third party system that Added them
         */
        [Ignore]
        public void TestFileWithContentTypeEntities()
        {
            Stream is1 = OpenXml4NetTestDataSamples.OpenSampleStream("ContentTypeHasEntities.ooxml");
            OPCPackage p = OPCPackage.Open(is1);

            // Check we found the contents of it
            bool foundCoreProps = false, foundDocument = false, foundTheme1 = false;
            foreach (PackagePart part in p.GetParts())
            {
                if (part.PartName.ToString().Equals("/docProps/core.xml"))
                {
                    Assert.AreEqual(ContentTypes.CORE_PROPERTIES_PART, part.ContentType);
                    foundCoreProps = true;
                }
                if (part.PartName.ToString().Equals("/word/document.xml"))
                {
                    Assert.AreEqual(XWPFRelation.DOCUMENT.ContentType, part.ContentType);
                    foundDocument = true;
                }
                if (part.PartName.ToString().Equals("/word/theme/theme1.xml"))
                {
                    Assert.AreEqual(XWPFRelation.THEME.ContentType, part.ContentType);
                    foundTheme1 = true;
                }
            }
            Assert.IsTrue(foundCoreProps, "Core not found in " + p.GetParts());
            Assert.IsTrue(foundDocument, "Document not found in " + p.GetParts());
            Assert.IsTrue(foundTheme1, "Theme1 not found in " + p.GetParts());
        }


        /**
         * Check that we can open a file where there are valid
         *  parameters on a content type
         */
        [Test]
        public void TestFileWithContentTypeParams()
        {
            Stream is1 = OpenXml4NetTestDataSamples.OpenSampleStream("ContentTypeHasParameters.ooxml");

            OPCPackage p = OPCPackage.Open(is1);

            String typeResqml = "application/x-resqml+xml";

            // Check the types on everything
            foreach (PackagePart part in p.GetParts())
            {
                // _rels type doesn't have any params
                if (part.IsRelationshipPart)
                {
                    Assert.AreEqual(ContentTypes.RELATIONSHIPS_PART, part.ContentType);
                    Assert.AreEqual(ContentTypes.RELATIONSHIPS_PART, part.ContentTypeDetails.ToString());
                    Assert.AreEqual(false, part.ContentTypeDetails.HasParameters());
                    Assert.AreEqual(0, part.ContentTypeDetails.GetParameterKeys().Length);
                }
                // Core type doesn't have any params
                else if (part.PartName.ToString().Equals("/docProps/core.xml"))
                {
                    Assert.AreEqual(ContentTypes.CORE_PROPERTIES_PART, part.ContentType);
                    Assert.AreEqual(ContentTypes.CORE_PROPERTIES_PART, part.ContentTypeDetails.ToString());
                    Assert.AreEqual(false, part.ContentTypeDetails.HasParameters());
                    Assert.AreEqual(0, part.ContentTypeDetails.GetParameterKeys().Length);
                }
                // Global Crs types do have params
                else if (part.PartName.ToString().Equals("/global1dCrs.xml"))
                {
                    Assert.AreEqual(typeResqml, part.ContentType.Substring(0, typeResqml.Length));
                    Assert.AreEqual(typeResqml, part.ContentTypeDetails.ToString(false));
                    Assert.AreEqual(true, part.ContentTypeDetails.HasParameters());
                    Assert.AreEqual(typeResqml + ";version=2.0;type=obj_global1dCrs", part.ContentTypeDetails.ToString());
                    Assert.AreEqual(2, part.ContentTypeDetails.GetParameterKeys().Length);
                    Assert.AreEqual("2.0", part.ContentTypeDetails.GetParameter("version"));
                    Assert.AreEqual("obj_global1dCrs", part.ContentTypeDetails.GetParameter("type"));
                }
                else if (part.PartName.ToString().Equals("/global2dCrs.xml"))
                {
                    Assert.AreEqual(typeResqml, part.ContentType.Substring(0, typeResqml.Length));
                    Assert.AreEqual(typeResqml, part.ContentTypeDetails.ToString(false));
                    Assert.AreEqual(true, part.ContentTypeDetails.HasParameters());
                    Assert.AreEqual(typeResqml + ";version=2.0;type=obj_global2dCrs", part.ContentTypeDetails.ToString());
                    Assert.AreEqual(2, part.ContentTypeDetails.GetParameterKeys().Length);
                    Assert.AreEqual("2.0", part.ContentTypeDetails.GetParameter("version"));
                    Assert.AreEqual("obj_global2dCrs", part.ContentTypeDetails.GetParameter("type"));
                }
                // Other thingy
                else if (part.PartName.ToString().Equals("/myTestingGuid.xml"))
                {
                    Assert.AreEqual(typeResqml, part.ContentType.Substring(0, typeResqml.Length));
                    Assert.AreEqual(typeResqml, part.ContentTypeDetails.ToString(false));
                    Assert.AreEqual(true, part.ContentTypeDetails.HasParameters());
                    Assert.AreEqual(typeResqml + ";version=2.0;type=obj_tectonicBoundaryFeature", part.ContentTypeDetails.ToString());
                    Assert.AreEqual(2, part.ContentTypeDetails.GetParameterKeys().Length);
                    Assert.AreEqual("2.0", part.ContentTypeDetails.GetParameter("version"));
                    Assert.AreEqual("obj_tectonicBoundaryFeature", part.ContentTypeDetails.GetParameter("type"));
                }
                // That should be it!
                else
                {
                    Assert.Fail("Unexpected part " + part);
                }
            }
        }
    }
}



