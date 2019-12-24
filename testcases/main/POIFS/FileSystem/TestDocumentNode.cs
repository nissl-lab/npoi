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


using System;
using System.Collections;
using System.IO;

using NUnit.Framework;

using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NPOI.POIFS.Storage;
using NPOI.POIFS.Properties;

namespace TestCases.POIFS.FileSystem
{
    /**
     * Class to Test DocumentNode functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestDocumentNode
    {

        /**
         * Constructor TestDocumentNode
         *
         * @param name
         */

        public TestDocumentNode()
        {

        }

        /**
         * Test constructor
         *
         * @exception IOException
         */
        [Test]
        public void TestConstructor()
        {
            DirectoryProperty property1 = new DirectoryProperty("directory");
            RawDataBlock[] rawBlocks = new RawDataBlock[4];
            MemoryStream stream =
                new MemoryStream(new byte[2048]);

            for (int j = 0; j < 4; j++)
            {
                rawBlocks[j] = new RawDataBlock(stream);
            }
            OPOIFSDocument document = new OPOIFSDocument("document", rawBlocks,
                                             2000);
            DocumentProperty property2 = document.DocumentProperty;
            DirectoryNode parent = new DirectoryNode(property1, (POIFSFileSystem)null, null);
            DocumentNode node = new DocumentNode(property2, parent);

            // Verify we can retrieve the document
            Assert.AreEqual(property2.Document, node.Document);

            // Verify we can Get the size
            Assert.AreEqual(property2.Size, node.Size);

            // Verify isDocumentEntry returns true
            Assert.IsTrue(node.IsDocumentEntry);

            // Verify isDirectoryEntry returns false
            Assert.IsTrue(!node.IsDirectoryEntry);

            // Verify GetName behaves correctly
            Assert.AreEqual(property2.Name, node.Name);

            // Verify GetParent behaves correctly
            Assert.AreEqual(parent, node.Parent);
        }
    }
}