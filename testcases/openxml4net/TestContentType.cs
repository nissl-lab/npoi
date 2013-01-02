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
         * Check rule [O1.2]: Format designers might restrict the usage of
         * parameters for content types.
         */
        [Test]
        public void TestContentTypeParameterFailure()
        {
            String[] contentTypesToTest = new String[] { "mail/toto;titi=tata",
				"text/xml;a=b;c=d", "mail/toto;\"titi=tata\"" };
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
    }
}



