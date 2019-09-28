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

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/


namespace TestCases.POIFS.Storage
{
    using System;
    using System.IO;
    using System.Collections;

    using NUnit.Framework;
    using NPOI.POIFS.Storage;
    using NPOI.POIFS.Common;
    using NPOI.Util;
    using NPOI.POIFS.Properties;
    using NPOI.POIFS.FileSystem;
    using System.Collections.Generic;

    /**
* Class to Test SmallBlockTableWriter functionality
*
* @author Marc Johnson
*/
    [TestFixture]
    public class TestSmallBlockTableWriter
    {

        /**
         * Constructor TestSmallBlockTableWriter
         *
         * @param name
         */

        public TestSmallBlockTableWriter()
        {
        }

        /**
         * Test writing constructor
         *
         * @exception IOException
         */
        [Test]
        public void TestWritingConstructor()
        {
            List<OPOIFSDocument> documents = new List<OPOIFSDocument>();

            documents.Add(
                new OPOIFSDocument(
                    "doc340", new MemoryStream(new byte[340])));
            documents.Add(
                new OPOIFSDocument(
                    "doc5000", new MemoryStream(new byte[5000])));
            documents
                .Add(new OPOIFSDocument("doc0",
                                       new MemoryStream(new byte[0])));
            documents
                .Add(new OPOIFSDocument("doc1",
                                       new MemoryStream(new byte[1])));
            documents
                .Add(new OPOIFSDocument("doc2",
                                       new MemoryStream(new byte[2])));
            documents
                .Add(new OPOIFSDocument("doc3",
                                       new MemoryStream(new byte[3])));
            documents
                .Add(new OPOIFSDocument("doc4",
                                       new MemoryStream(new byte[4])));
            documents
                .Add(new OPOIFSDocument("doc5",
                                       new MemoryStream(new byte[5])));
            documents
                .Add(new OPOIFSDocument("doc6",
                                       new MemoryStream(new byte[6])));
            documents
                .Add(new OPOIFSDocument("doc7",
                                       new MemoryStream(new byte[7])));
            documents
                .Add(new OPOIFSDocument("doc8",
                                       new MemoryStream(new byte[8])));
            documents
                .Add(new OPOIFSDocument("doc9",
                                       new MemoryStream(new byte[9])));
            HeaderBlock header = new HeaderBlock(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            RootProperty root = new PropertyTable(header).Root;
            SmallBlockTableWriter sbtw = new SmallBlockTableWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, documents, root);
            BlockAllocationTableWriter bat = sbtw.SBAT;

            // 15 small blocks: 6 for doc340, 0 for doc5000 (too big), 0
            // for doc0 (no storage needed), 1 each for doc1 through doc9
            Assert.AreEqual(15 * 64, root.Size);

            // 15 small blocks rounds up to 2 big blocks
            Assert.AreEqual(2, sbtw.CountBlocks);
            int start_block = 1000 + root.StartBlock;

            sbtw.StartBlock = start_block;
            Assert.AreEqual(start_block, root.StartBlock);
        }
    }
}