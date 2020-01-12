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
    using NUnit.Framework;
    using System;

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
    public class TestPackageCorePropertiesGetKeywords
    {
        [Test]
        public void TestGetSetKeywords()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("TestPoiXMLDocumentCorePropertiesGetKeywords.docx");
            String keywords = doc.GetProperties().CoreProperties.Keywords;
            Assert.AreEqual("extractor, test, rdf", keywords);

            doc.GetProperties().CoreProperties.Keywords =  ("test, keywords");
            doc = XWPFTestDataSamples.WriteOutAndReadBack(doc);
            keywords = doc.GetProperties().CoreProperties.Keywords;
            Assert.AreEqual("test, keywords", keywords);
        }
    }

}