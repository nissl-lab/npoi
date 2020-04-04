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
 * Leon Wang
 * ==============================================================*/



using System;
using System.Collections;
using System.IO;

using NUnit.Framework;

using NPOI.POIFS.FileSystem;
using NPOI.POIFS.Properties;
using NPOI.POIFS.Storage;
using System.Collections.Generic;
/**
* Class to Test DirectoryNode functionality
*
* @author Marc Johnson
*/

namespace TestCases.POIFS.FileSystem
{
    /// <summary>
    /// Summary description for TestDirectoryNode
    /// </summary>
    [TestFixture]
    public class TestDirectoryNode
    {

        /**
         * Constructor TestDirectoryNode
         *
         * @param name
         */

        public TestDirectoryNode()
        {

        }

        /**
         * Test trivial constructor (a DirectoryNode with no children)
         *
         * @exception IOException
         */
        [Test]
        public void TestEmptyConstructor()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            DirectoryProperty property1 = new DirectoryProperty("parent");
            DirectoryProperty property2 = new DirectoryProperty("child");
            DirectoryNode parent = new DirectoryNode(property1, fs, null);
            DirectoryNode node = new DirectoryNode(property2, fs, parent);

            Assert.AreEqual(0, parent.Path.Length);
            Assert.AreEqual(1, node.Path.Length);
            Assert.AreEqual("child", node.Path.GetComponent(0));

            // Verify that GetEntries behaves correctly
            int count = 0;
            IEnumerator<Entry> iter = node.Entries;

            while (iter.MoveNext())
            {
                count++;
            }
            Assert.AreEqual(0, count);

            // Verify behavior of IsEmpty
            Assert.IsTrue(node.IsEmpty);

            // Verify behavior of EntryCount
            Assert.AreEqual(0, node.EntryCount);

            // Verify behavior of Entry
            try
            {
                node.GetEntry("foo");
                Assert.Fail("Should have caught FileNotFoundException");
            }
            catch (FileNotFoundException )
            {

                // as expected
            }

            // Verify behavior of isDirectoryEntry
            Assert.IsTrue(node.IsDirectoryEntry);

            // Verify behavior of GetName
            Assert.AreEqual(property2.Name, node.Name);

            // Verify behavior of isDocumentEntry
            Assert.IsTrue(!node.IsDocumentEntry);

            // Verify behavior of GetParent
            Assert.AreEqual(parent, node.Parent);
        }

        /**
         * Test non-trivial constructor (a DirectoryNode with children)
         *
         * @exception IOException
         */
        [Test]
        public void TestNonEmptyConstructor()
        {
            DirectoryProperty property1 = new DirectoryProperty("parent");
            DirectoryProperty property2 = new DirectoryProperty("child1");

            property1.AddChild(property2);
            property1.AddChild(new DocumentProperty("child2", 2000));
            property2.AddChild(new DocumentProperty("child3", 30000));
            DirectoryNode node = new DirectoryNode(property1, new POIFSFileSystem(), null);

            // Verify that GetEntries behaves correctly
            int count = 0;
            IEnumerator<Entry> iter = node.Entries;

            while (iter.MoveNext())
            {
                count++;
                //iter.Current;
            }
            Assert.AreEqual(2, count);

            // Verify behavior of IsEmpty
            Assert.IsTrue(!node.IsEmpty);

            // Verify behavior of EntryCount
            Assert.AreEqual(2, node.EntryCount);

            // Verify behavior of Entry
            DirectoryNode child1 = (DirectoryNode)node.GetEntry("child1");

            child1.GetEntry("child3");
            node.GetEntry("child2");
            try
            {
                node.GetEntry("child3");
                Assert.Fail("Should have caught FileNotFoundException");
            }
            catch (FileNotFoundException)
            {

                // as expected
            }

            // Verify behavior of isDirectoryEntry
            Assert.IsTrue(node.IsDirectoryEntry);

            // Verify behavior of GetName
            Assert.AreEqual(property1.Name, node.Name);

            // Verify behavior of isDocumentEntry
            Assert.IsTrue(!node.IsDocumentEntry);

            // Verify behavior of GetParent
            Assert.IsNull(node.Parent);
        }

        /**
         * Test deletion methods
         *
         * @exception IOException
         */
        [Test]
        public void TestDeletion()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            DirectoryEntry root = fs.Root;

            Assert.IsFalse(root.Delete());
            Assert.IsTrue(root.IsEmpty);

            DirectoryEntry dir = fs.CreateDirectory("myDir");

            Assert.IsFalse(root.IsEmpty);
            Assert.IsTrue(dir.IsEmpty);

            Assert.IsFalse(root.Delete());
            // Verify can Delete empty directory
            Assert.IsTrue(dir.Delete());
            dir = fs.CreateDirectory("NextDir");
            DocumentEntry doc = dir.CreateDocument("foo", new MemoryStream(new byte[1]));

            Assert.IsFalse(root.IsEmpty);
            Assert.IsFalse(dir.IsEmpty);

            Assert.IsFalse(dir.Delete());

            // Verify cannot Delete empty directory
            Assert.IsTrue(!dir.Delete());
            Assert.IsTrue(doc.Delete());
            Assert.IsTrue(dir.IsEmpty);
            // Verify now we can Delete it
            Assert.IsTrue(dir.Delete());
            Assert.IsTrue(root.IsEmpty);

            fs.Close();
        }

        /**
         * Test Change name methods
         *
         * @exception IOException
         */
        [Test]
        public void TestRename()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            DirectoryEntry root = fs.Root;

            // Verify cannot Rename the root directory
            Assert.IsTrue(!root.RenameTo("foo"));
            DirectoryEntry dir = fs.CreateDirectory("myDir");

            Assert.IsTrue(dir.RenameTo("foo"));
            Assert.AreEqual("foo", dir.Name);
            DirectoryEntry dir2 = fs.CreateDirectory("myDir");

            Assert.IsTrue(!dir2.RenameTo("foo"));
            Assert.AreEqual("myDir", dir2.Name);
            Assert.IsTrue(dir.RenameTo("FirstDir"));
            Assert.IsTrue(dir2.RenameTo("foo"));
            Assert.AreEqual("foo", dir2.Name);

            fs.Close();
        }

    }
}