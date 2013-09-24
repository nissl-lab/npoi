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


namespace TestCases.HWPF.Model
{
    
    using NPOI.HWPF.Model;
    using NPOI.HWPF;
    using NUnit.Framework;
    /**
     * Test the table which handles author revision marks
     */
    [TestFixture]
    public class TestRevisionMarkAuthorTable
    {
        /**
         * Tests that an empty file doesn't have one
         */
        [Test]
        public void TestEmptyDocument()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("empty.doc");

            RevisionMarkAuthorTable rmt = doc.GetRevisionMarkAuthorTable();
            Assert.IsNull(rmt);
        }

        /**
         * Tests that we can load a document with
         *  only simple entries in the table
         */
        [Test]
        public void TestSimpleDocument()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("two_images.doc");

            RevisionMarkAuthorTable rmt = doc.GetRevisionMarkAuthorTable();
            Assert.IsNotNull(rmt);
            Assert.AreEqual(1, rmt.GetSize());
            Assert.AreEqual("Unknown", rmt.GetAuthor(0));

            Assert.AreEqual(null, rmt.GetAuthor(1));
            Assert.AreEqual(null, rmt.GetAuthor(2));
            Assert.AreEqual(null, rmt.GetAuthor(3));
        }

        /**
         * Several authors, one of whom has no name
         */
        [Test]
        public void TestMultipleAuthors()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("MarkAuthorsTable.doc");

            RevisionMarkAuthorTable rmt = doc.GetRevisionMarkAuthorTable();
            Assert.IsNotNull(rmt);
            Assert.AreEqual(4, rmt.GetSize());
            Assert.AreEqual("Unknown", rmt.GetAuthor(0));
            Assert.AreEqual("BSanders", rmt.GetAuthor(1));
            Assert.AreEqual(" ", rmt.GetAuthor(2));
            Assert.AreEqual("Ryan Lauck", rmt.GetAuthor(3));

            Assert.AreEqual(null, rmt.GetAuthor(4));
            Assert.AreEqual(null, rmt.GetAuthor(5));
            Assert.AreEqual(null, rmt.GetAuthor(6));
        }
    }
}

