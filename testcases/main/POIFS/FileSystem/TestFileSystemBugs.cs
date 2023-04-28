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

namespace TestCases.POIFS.FileSystem
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using NPOI.POIFS.FileSystem;
    using NUnit.Framework;
    using TestCases;

    /**
     * Tests bugs across both POIFSFileSystem and NPOIFSFileSystem
     */
    [TestFixture]
    public class TestFileSystemBugs
    {
        protected static POIDataSamples _samples = POIDataSamples.GetPOIFSInstance();
        protected static POIDataSamples _ssSamples = POIDataSamples.GetSpreadSheetInstance();

        protected List<NPOIFSFileSystem> openedFSs;
        protected void tearDown()
        {
            if (openedFSs != null && openedFSs.Count > 0)
            {
                foreach (NPOIFSFileSystem fs in openedFSs)
                {
                    try
                    {
                        if (fs != null)
                            fs.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error closing FS: " + e);
                    }
                }
            }
            openedFSs = null;
        }
        protected DirectoryNode[] openSample(String name, bool oldFails)
        {
            return openSamples(new Stream[] {
                _samples.OpenResourceAsStream(name),
                _samples.OpenResourceAsStream(name)
        }, oldFails);
        }
        protected DirectoryNode[] openSSSample(String name, bool oldFails)
        {
            return openSamples(new Stream[] {
                _ssSamples.OpenResourceAsStream(name),
                _ssSamples.OpenResourceAsStream(name)
        }, oldFails);
        }
        protected DirectoryNode[] openSamples(Stream[] inps, bool oldFails)
        {
            NPOIFSFileSystem nfs = new NPOIFSFileSystem(inps[0]);
            if (openedFSs == null)
                openedFSs = new List<NPOIFSFileSystem>();
            openedFSs.Add(nfs);

            OPOIFSFileSystem ofs = null;
            try
            {
                ofs = new OPOIFSFileSystem(inps[1]);
                if (oldFails)
                    Assert.Fail("POIFSFileSystem should have failed but didn't");
            }
            catch
            {
                if (!oldFails)
                    throw;
            }

            if (ofs == null)
                return new DirectoryNode[] { nfs.Root };
            return new DirectoryNode[] { ofs.Root, nfs.Root };
        }

        /**
         * Test that we can open files that come via Lotus notes.
         * These have a top level directory without a name....
         */
        [Test]
        public void TestNotesOLE2Files()
        {
            // Check the contents
            foreach (DirectoryNode root in openSample("Notes.ole2", false))
            {
                Assert.AreEqual(1, root.EntryCount);
                IEnumerator<Entry> it = root.Entries;
                it.MoveNext();
                Entry entry = it.Current;

                Assert.IsTrue(entry.IsDirectoryEntry);
                Assert.IsTrue(entry is DirectoryEntry);

                // The directory lacks a name!
                DirectoryEntry dir = (DirectoryEntry)entry;
                Assert.AreEqual("", dir.Name);

                // Has two children
                Assert.AreEqual(2, dir.EntryCount);

                it = dir.Entries;
                // Check them
                it.MoveNext();
                entry = it.Current;
                Assert.AreEqual(true, entry.IsDocumentEntry);
                Assert.AreEqual("\u0001Ole10Native", entry.Name);

                it.MoveNext();
                entry = it.Current;
                Assert.AreEqual(true, entry.IsDocumentEntry);
                Assert.AreEqual("\u0001CompObj", entry.Name);
            }
        }

        /**
         * Ensure that a file with a corrupted property in the
         *  properties table can still be loaded, and the remaining
         *  properties used
         * Note - only works for NPOIFSFileSystem, POIFSFileSystem
         *  can't cope with this level of corruption
         */
        [Test]
        public void TestCorruptedProperties()
        {
            foreach (DirectoryNode root in openSample("unknown_properties.msg", true))
            {
                Assert.AreEqual(42, root.EntryCount);
            }
        }

        /**
         * With heavily nested documents, ensure we still re-write the same
         */
        [Test]
        public void TestHeavilyNestedReWrite()
        {
            foreach (DirectoryNode root in openSSSample("ex42570-20305.xls", false))
            {
                // Record the structure
                Dictionary<String, int> entries = new Dictionary<String, int>();
                fetchSizes("/", root, entries);

                // Prepare to copy
                DirectoryNode dest;
                if (root.NFileSystem != null)
                {
                    dest = (new NPOIFSFileSystem()).Root;
                }
                else
                {
                    dest = (new OPOIFSFileSystem()).Root;
                }

                // Copy over
                EntryUtils.CopyNodes(root, dest);

                // Re-load, always as NPOIFS
                MemoryStream baos = new MemoryStream();
                if (root.NFileSystem != null)
                {
                    root.NFileSystem.WriteFileSystem(baos);
                }
                else
                {
                    root.OFileSystem.WriteFileSystem(baos);
                }
                NPOIFSFileSystem read = new NPOIFSFileSystem(
                        new MemoryStream(baos.ToArray()));

                // Check the structure matches
                CheckSizes("/", read.Root, entries);
            }
        }
        private void fetchSizes(String path, DirectoryNode dir, Dictionary<String, int> entries)
        {
            foreach (Entry entry in dir)
            {
                if (entry is DirectoryNode)
                {
                    String ourPath = path + entry.Name + "/";
                    entries.Add(ourPath, -1);
                    fetchSizes(ourPath, (DirectoryNode)entry, entries);
                }
                else
                {
                    DocumentNode doc = (DocumentNode)entry;
                    entries.Add(path + entry.Name, doc.Size);
                }
            }
        }
        private void CheckSizes(String path, DirectoryNode dir, Dictionary<String, int> entries)
        {
            foreach (Entry entry in dir)
            {
                if (entry is DirectoryNode)
                {
                    String ourPath = path + entry.Name + "/";
                    Assert.IsTrue(entries.ContainsKey(ourPath));
                    Assert.AreEqual(-1, entries[(ourPath)]);
                    CheckSizes(ourPath, (DirectoryNode)entry, entries);
                }
                else
                {
                    DocumentNode doc = (DocumentNode)entry;
                    Assert.AreEqual(entries[(path + entry.Name)], doc.Size);
                }
            }
        }
    }

}