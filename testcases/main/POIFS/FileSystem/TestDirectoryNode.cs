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

using NUnit.Framework;using NUnit.Framework.Legacy;

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

            ClassicAssert.AreEqual(0, parent.Path.Length);
            ClassicAssert.AreEqual(1, node.Path.Length);
            ClassicAssert.AreEqual("child", node.Path.GetComponent(0));

            // Verify that GetEntries behaves correctly
            int count = 0;
            IEnumerator<Entry> iter = node.Entries;

            while (iter.MoveNext())
            {
                count++;
            }
            ClassicAssert.AreEqual(0, count);

            // Verify behavior of IsEmpty
            ClassicAssert.IsTrue(node.IsEmpty);

            // Verify behavior of EntryCount
            ClassicAssert.AreEqual(0, node.EntryCount);

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
            ClassicAssert.IsTrue(node.IsDirectoryEntry);

            // Verify behavior of GetName
            ClassicAssert.AreEqual(property2.Name, node.Name);

            // Verify behavior of isDocumentEntry
            ClassicAssert.IsTrue(!node.IsDocumentEntry);

            // Verify behavior of GetParent
            ClassicAssert.AreEqual(parent, node.Parent);
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
            ClassicAssert.AreEqual(2, count);

            // Verify behavior of IsEmpty
            ClassicAssert.IsTrue(!node.IsEmpty);

            // Verify behavior of EntryCount
            ClassicAssert.AreEqual(2, node.EntryCount);

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
            ClassicAssert.IsTrue(node.IsDirectoryEntry);

            // Verify behavior of GetName
            ClassicAssert.AreEqual(property1.Name, node.Name);

            // Verify behavior of isDocumentEntry
            ClassicAssert.IsTrue(!node.IsDocumentEntry);

            // Verify behavior of GetParent
            ClassicAssert.IsNull(node.Parent);
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

            ClassicAssert.IsFalse(root.Delete());
            ClassicAssert.IsTrue(root.IsEmpty);

            DirectoryEntry dir = fs.CreateDirectory("myDir");

            ClassicAssert.IsFalse(root.IsEmpty);
            ClassicAssert.IsTrue(dir.IsEmpty);

            ClassicAssert.IsFalse(root.Delete());
            // Verify can Delete empty directory
            ClassicAssert.IsTrue(dir.Delete());
            dir = fs.CreateDirectory("NextDir");
            DocumentEntry doc = dir.CreateDocument("foo", new MemoryStream(new byte[1]));

            ClassicAssert.IsFalse(root.IsEmpty);
            ClassicAssert.IsFalse(dir.IsEmpty);

            ClassicAssert.IsFalse(dir.Delete());

            // Verify cannot Delete empty directory
            ClassicAssert.IsTrue(!dir.Delete());
            ClassicAssert.IsTrue(doc.Delete());
            ClassicAssert.IsTrue(dir.IsEmpty);
            // Verify now we can Delete it
            ClassicAssert.IsTrue(dir.Delete());
            ClassicAssert.IsTrue(root.IsEmpty);

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
            ClassicAssert.IsTrue(!root.RenameTo("foo"));
            DirectoryEntry dir = fs.CreateDirectory("myDir");

            ClassicAssert.IsTrue(dir.RenameTo("foo"));
            ClassicAssert.AreEqual("foo", dir.Name);
            DirectoryEntry dir2 = fs.CreateDirectory("myDir");

            ClassicAssert.IsTrue(!dir2.RenameTo("foo"));
            ClassicAssert.AreEqual("myDir", dir2.Name);
            ClassicAssert.IsTrue(dir.RenameTo("FirstDir"));
            ClassicAssert.IsTrue(dir2.RenameTo("foo"));
            ClassicAssert.AreEqual("foo", dir2.Name);

            fs.Close();
        }

    }
}