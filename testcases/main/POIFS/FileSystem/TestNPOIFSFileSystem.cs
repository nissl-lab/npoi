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

using System;
using System.IO;
using System.Collections.Generic;

using NUnit.Framework;using NUnit.Framework.Legacy;
using NPOI.POIFS.Storage;
using NPOI.POIFS.Properties;
using NPOI.POIFS.FileSystem;
using NPOI.POIFS.Common;
using NPOI.HPSF;
using NPOI.POIFS.NIO;
using NPOI.Util;
using NPOI.POIFS.EventFileSystem;
using NUnit.Framework.Constraints;
using Property = NPOI.POIFS.Properties.Property;

namespace TestCases.POIFS.FileSystem
{
    /// <summary>
    /// Summary description for TestNPOIFSFileSystem
    /// </summary>
    [TestFixture]
    public class TestNPOIFSFileSystem
    {
        private static POIDataSamples _inst = POIDataSamples.GetPOIFSInstance();
        public TestNPOIFSFileSystem()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        /**
        * Returns test files with 512 byte and 4k block sizes, loaded
        *  both from InputStreams and Files
        */
        protected NPOIFSFileSystem[] get512and4kFileAndInput()
        {
            NPOIFSFileSystem fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));
            NPOIFSFileSystem fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));
            NPOIFSFileSystem fsC = new NPOIFSFileSystem(_inst.GetFile("BlockSize4096.zvi"));
            NPOIFSFileSystem fsD = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize4096.zvi"));
            return new NPOIFSFileSystem[] { fsA, fsB, fsC, fsD };
        }

        protected static void assertBATCount(NPOIFSFileSystem fs, int expectedBAT, int expectedXBAT)
        {
            int foundBAT = 0;
            int foundXBAT = 0;
            int sz = (int)(fs.Size / fs.GetBigBlockSize());
            for (int i = 0; i < sz; i++)
            {
                if (fs.GetNextBlock(i) == POIFSConstants.FAT_SECTOR_BLOCK)
                {
                    foundBAT++;
                }
                if (fs.GetNextBlock(i) == POIFSConstants.DIFAT_SECTOR_BLOCK)
                {
                    foundXBAT++;
                }
            }
            ClassicAssert.AreEqual(expectedBAT, foundBAT, "Wrong number of BATs");
            ClassicAssert.AreEqual(expectedXBAT, foundXBAT, "Wrong number of XBATs with " + expectedBAT + " BATs");
        }
        protected void assertContentsMatches(byte[] expected, DocumentEntry doc)
        {
            NDocumentInputStream inp = new NDocumentInputStream(doc);
            byte[] contents = new byte[doc.Size];
            ClassicAssert.AreEqual(doc.Size, inp.Read(contents));
            inp.Close();

            if (expected != null)
                Assert.That(expected, new EqualConstraint(contents));
        }
        public static HeaderBlock WriteOutAndReadHeader(NPOIFSFileSystem fs)
        {
            MemoryStream baos = new MemoryStream();
            fs.WriteFileSystem(baos);

            HeaderBlock header = new HeaderBlock(new MemoryStream(baos.ToArray()));
            return header;
        }

        public static NPOIFSFileSystem WriteOutAndReadBack(NPOIFSFileSystem original)
        {
            MemoryStream baos = new MemoryStream();
            original.WriteFileSystem(baos);
            original.Close();
            return new NPOIFSFileSystem(new ByteArrayInputStream(baos.ToArray()));
        }

        protected static NPOIFSFileSystem WriteOutFileAndReadBack(NPOIFSFileSystem original)
        {
            FileInfo file = TempFile.CreateTempFile("TestPOIFS", ".ole2");
            try
            {
                using(FileStream fout = file.Open(FileMode.OpenOrCreate, FileAccess.ReadWrite))
                { 
                    original.WriteFileSystem(fout);
                }
            }
            catch
            {
                // avoid throw exception
            }
            original.Close();
            return new NPOIFSFileSystem(file, false);
        }

        [Test]
        public void TestBasicOpen()
        {
            NPOIFSFileSystem fsA, fsB;
            fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));
            fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));

            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB })
            {
                ClassicAssert.AreEqual(512, fs.GetBigBlockSize());
            }

            fsA.Close();
            fsB.Close();

            fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize4096.zvi"));
            fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize4096.zvi"));

            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB })
            {
                ClassicAssert.AreEqual(4096, fs.GetBigBlockSize());
            }
            fsA.Close();
            fsB.Close();
        }

        [Test]
        public void TestPropertiesAndFatOnRead()
        {
            NPOIFSFileSystem fsA, fsB;

            fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));
            fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));
            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB })
            {
                fs.GetBATBlockAndIndex(0);
                fs.GetBATBlockAndIndex(1);

                try
                {
                    fs.GetBATBlockAndIndex(140);
                    Assert.Fail("Should only be one BAT, but a 2nd was found");
                }
                //catch (IndexOutOfRangeException)
                catch (ArgumentOutOfRangeException)
                {

                }

                ClassicAssert.AreEqual(98, fs.GetNextBlock(97));
                ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(98));

                NPropertyTable props = fs.PropertyTable;
                ClassicAssert.AreEqual(90, props.StartBlock);
                ClassicAssert.AreEqual(7, props.CountBlocks);

                RootProperty root = props.Root;
                ClassicAssert.AreEqual("Root Entry", root.Name);
                ClassicAssert.AreEqual(11564, root.Size);
                ClassicAssert.AreEqual(0, root.StartBlock);

                NPOI.POIFS.Properties.Property prop;
                IEnumerator<NPOI.POIFS.Properties.Property> pi = root.Children;
                //prop = pi.Current;
                pi.MoveNext();
                prop = pi.Current;
                ClassicAssert.AreEqual("Thumbnail", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                ClassicAssert.AreEqual("\x0005DocumentSummaryInformation", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                ClassicAssert.AreEqual("\x0005SummaryInformation", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                ClassicAssert.AreEqual("Image", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                ClassicAssert.AreEqual(false, pi.MoveNext());

                NPOIFSMiniStore miniStore = fs.GetMiniStore();
                miniStore.GetBATBlockAndIndex(0);
                miniStore.GetBATBlockAndIndex(128);

                try
                {
                    miniStore.GetBATBlockAndIndex(256);
                    Assert.Fail("Should only be two SBATs, but a 3rd was found");
                }
                //catch (IndexOutOfRangeException)
                catch (ArgumentOutOfRangeException)
                {
                }

                for (int i = 0; i < 50; i++)
                    ClassicAssert.AreEqual(i + 1, miniStore.GetNextBlock(i));

                ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, miniStore.GetNextBlock(50));

                fs.Close();
            }
            fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize4096.zvi"));
            fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize4096.zvi"));
            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB })
            {
                fs.GetBATBlockAndIndex(0);
                fs.GetBATBlockAndIndex(1);

                try
                {
                    fs.GetBATBlockAndIndex(1040);
                    Assert.Fail("Should only be one BAT, but a 2nd was found");
                }
                //catch (IndexOutOfRangeException)
                catch (ArgumentOutOfRangeException)
                {
                }

                ClassicAssert.AreEqual(1, fs.GetNextBlock(0));
                ClassicAssert.AreEqual(2, fs.GetNextBlock(1));
                ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(2));

                NPropertyTable props = fs.PropertyTable;
                ClassicAssert.AreEqual(12, props.StartBlock);
                ClassicAssert.AreEqual(1, props.CountBlocks);

                RootProperty root = props.Root;
                ClassicAssert.AreEqual("Root Entry", root.Name);
                ClassicAssert.AreEqual(11564, root.Size);
                ClassicAssert.AreEqual(0, root.StartBlock);

                NPOI.POIFS.Properties.Property prop;
                IEnumerator<NPOI.POIFS.Properties.Property> pi = root.Children;

                pi.MoveNext();
                prop = pi.Current;
                ClassicAssert.AreEqual("Thumbnail", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                ClassicAssert.AreEqual("\x0005DocumentSummaryInformation", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                ClassicAssert.AreEqual("\x0005SummaryInformation", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                ClassicAssert.AreEqual("Image", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                ClassicAssert.AreEqual("Tags", prop.Name);
                ClassicAssert.AreEqual(false, pi.MoveNext());

                NPOIFSMiniStore miniStore = fs.GetMiniStore();
                miniStore.GetBATBlockAndIndex(0);
                miniStore.GetBATBlockAndIndex(128);
                miniStore.GetBATBlockAndIndex(1023);

                try
                {
                    miniStore.GetBATBlockAndIndex(1024);
                    Assert.Fail("Should only be one SBAT, but a 2nd was found");
                }
                //catch(IndexOutOfRangeException)
                catch (ArgumentOutOfRangeException)
                {
                }

                for (int i = 0; i < 50; i++)
                    ClassicAssert.AreEqual(i + 1, miniStore.GetNextBlock(i));

                ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, miniStore.GetNextBlock(50));
                fs.Close();
            }
        }

        [Test]
        public void TestNextBlock()
        {
            NPOIFSFileSystem fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));
            NPOIFSFileSystem fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));

            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB })
            {
                for (int i = 0; i < 21; i++)
                    ClassicAssert.AreEqual(i + 1, fs.GetNextBlock(i));

                // 21 jumps to 89, then ends
                ClassicAssert.AreEqual(89, fs.GetNextBlock(21));
                ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(89));

                for (int i = 22; i < 88; i++)
                    ClassicAssert.AreEqual(i + 1, fs.GetNextBlock(i));
                ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(88));

                for (int i = 90; i < 96; i++)
                {
                    ClassicAssert.AreEqual(i + 1, fs.GetNextBlock(i));
                }
                ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(96));

                ClassicAssert.AreEqual(98, fs.GetNextBlock(97));
                ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(98));

                ClassicAssert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(99));

                //Leon i = 100
                for (int i = 100; i < fs.GetBigBlockSizeDetails().GetBATEntriesPerBlock(); i++)
                    ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(i));

                fs.Close();
            }

            fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize4096.zvi"));
            fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize4096.zvi"));

            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB })
            {
                ClassicAssert.AreEqual(1, fs.GetNextBlock(0));
                ClassicAssert.AreEqual(2, fs.GetNextBlock(1));
                ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(2));

                for (int i = 4; i < 11; i++)
                    ClassicAssert.AreEqual(i + 1, fs.GetNextBlock(i));

                ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(11));
                fs.Close();
            }

        }

        [Test]
        public void TestGetBlock()
        {
            NPOIFSFileSystem fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));
            NPOIFSFileSystem fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));

            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB })
            {
                ByteBuffer b;

                b = fs.GetBlockAt(0);
                ClassicAssert.AreEqual((byte)0x9e, b.Read());
                ClassicAssert.AreEqual((byte)0x75, b.Read());
                ClassicAssert.AreEqual((byte)0x97, b.Read());
                ClassicAssert.AreEqual((byte)0xf6, b.Read());

                b = fs.GetBlockAt(1);
                ClassicAssert.AreEqual((byte)0x86, b.Read());
                ClassicAssert.AreEqual((byte)0x09, b.Read());
                ClassicAssert.AreEqual((byte)0x22, b.Read());
                ClassicAssert.AreEqual((byte)0xfb, b.Read());

                b = fs.GetBlockAt(99);
                ClassicAssert.AreEqual((byte)0x01, b.Read());
                ClassicAssert.AreEqual((byte)0x00, b.Read());
                ClassicAssert.AreEqual((byte)0x00, b.Read());
                ClassicAssert.AreEqual((byte)0x00, b.Read());
                ClassicAssert.AreEqual((byte)0x02, b.Read());
                ClassicAssert.AreEqual((byte)0x00, b.Read());
                ClassicAssert.AreEqual((byte)0x00, b.Read());
                ClassicAssert.AreEqual((byte)0x00, b.Read());
                fs.Close();
            }

            fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize4096.zvi"));
            fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize4096.zvi"));


            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB })
            {
                ByteBuffer b;

                b = fs.GetBlockAt(0);
                ClassicAssert.AreEqual((byte)0x9e, b.Read());
                ClassicAssert.AreEqual((byte)0x75, b.Read());
                ClassicAssert.AreEqual((byte)0x97, b.Read());
                ClassicAssert.AreEqual((byte)0xf6, b.Read());

                b = fs.GetBlockAt(1);
                ClassicAssert.AreEqual((byte)0x00, b.Read());
                ClassicAssert.AreEqual((byte)0x00, b.Read());
                ClassicAssert.AreEqual((byte)0x03, b.Read());
                ClassicAssert.AreEqual((byte)0x00, b.Read());

                b = fs.GetBlockAt(14);
                ClassicAssert.AreEqual((byte)0x01, b.Read());
                ClassicAssert.AreEqual((byte)0x00, b.Read());
                ClassicAssert.AreEqual((byte)0x00, b.Read());
                ClassicAssert.AreEqual((byte)0x00, b.Read());
                ClassicAssert.AreEqual((byte)0x02, b.Read());
                ClassicAssert.AreEqual((byte)0x00, b.Read());
                ClassicAssert.AreEqual((byte)0x00, b.Read());
                ClassicAssert.AreEqual((byte)0x00, b.Read());

                fs.Close();
            }
        }

        [Test]
        public void TestGetFreeBlockWithSpare()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));

            ClassicAssert.AreEqual(true, fs.GetBATBlockAndIndex(0).Block.HasFreeSectors);

            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(100));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(101));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(102));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(103));

            ClassicAssert.AreEqual(100, fs.GetFreeBlock());

            ClassicAssert.AreEqual(100, fs.GetFreeBlock());

            fs.SetNextBlock(100, POIFSConstants.END_OF_CHAIN);
            ClassicAssert.AreEqual(101, fs.GetFreeBlock());

            fs.Close();
        }

        [Test]
        public void TestGetFreeBlockWithNoneSpare()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));
            int free;

            ClassicAssert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(99));
            assertBATCount(fs, 1, 0);
            for (int i = 100; i < 128; i++)
                ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(i));

            ClassicAssert.AreEqual(true, fs.GetBATBlockAndIndex(0).Block.HasFreeSectors);

            for (int i = 100; i < 128; i++)
                fs.SetNextBlock(i, POIFSConstants.END_OF_CHAIN);

            ClassicAssert.AreEqual(false, fs.GetBATBlockAndIndex(0).Block.HasFreeSectors);

            try
            {
                ClassicAssert.AreEqual(false, fs.GetBATBlockAndIndex(128).Block.HasFreeSectors);
                Assert.Fail("Should only be one BAT");
            }
            //catch (IndexOutOfRangeException)
            catch (ArgumentOutOfRangeException)
            {
            }
            assertBATCount(fs, 1, 0);

            // Now ask for a free one, will need to extend the file

            ClassicAssert.AreEqual(129, fs.GetFreeBlock());

            ClassicAssert.AreEqual(false, fs.GetBATBlockAndIndex(0).Block.HasFreeSectors);
            ClassicAssert.AreEqual(true, fs.GetBATBlockAndIndex(128).Block.HasFreeSectors);
            ClassicAssert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(128));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(129));

            // We now have 2 BATs, but no XBATs
            assertBATCount(fs, 2, 0);

            // Fill up to hold 109 BAT blocks
            for (int i = 0; i < 109; i++)
            {
                fs.GetFreeBlock();
                int startOffset = i * 128;
                while (fs.GetBATBlockAndIndex(startOffset).Block.HasFreeSectors)
                {
                    free = fs.GetFreeBlock();
                    fs.SetNextBlock(free, POIFSConstants.END_OF_CHAIN);
                }
            }

            ClassicAssert.AreEqual(false, fs.GetBATBlockAndIndex(109 * 128 - 1).Block.HasFreeSectors);
            try
            {
                ClassicAssert.AreEqual(false, fs.GetBATBlockAndIndex(109 * 128).Block.HasFreeSectors);
                Assert.Fail("Should only be 109 BATs");
            }
            // catch (IndexOutOfRangeException)
            catch (ArgumentOutOfRangeException)
            {
            }

            // We now have 109 BATs, but no XBATs
            assertBATCount(fs, 109, 0);


            // Ask for it to be written out, and check the header
            HeaderBlock header = WriteOutAndReadHeader(fs);
            ClassicAssert.AreEqual(109, header.BATCount);
            ClassicAssert.AreEqual(0, header.XBATCount);

            free = fs.GetFreeBlock();
            ClassicAssert.AreEqual(false, fs.GetBATBlockAndIndex(109 * 128 - 1).Block.HasFreeSectors);
            ClassicAssert.AreEqual(true, fs.GetBATBlockAndIndex(110 * 128 - 1).Block.HasFreeSectors);
            try
            {
                ClassicAssert.AreEqual(false, fs.GetBATBlockAndIndex(110 * 128).Block.HasFreeSectors);
                Assert.Fail("Should only be 110 BATs");
            }
            //catch (IndexOutOfRangeException)
            catch (ArgumentOutOfRangeException)
            {
            }

            assertBATCount(fs, 110, 1);

            header = WriteOutAndReadHeader(fs);
            ClassicAssert.AreEqual(110, header.BATCount);
            ClassicAssert.AreEqual(1, header.XBATCount);

            for (int i = 109; i < 109 + 127; i++)
            {
                fs.GetFreeBlock();
                int startOffset = i * 128;
                while (fs.GetBATBlockAndIndex(startOffset).Block.HasFreeSectors)
                {
                    free = fs.GetFreeBlock();
                    fs.SetNextBlock(free, POIFSConstants.END_OF_CHAIN);
                }

                assertBATCount(fs, i + 1, 1);
            }

            // Should now have 109+127 = 236 BATs
            ClassicAssert.AreEqual(false, fs.GetBATBlockAndIndex(236 * 128 - 1).Block.HasFreeSectors);
            try
            {
                ClassicAssert.AreEqual(false, fs.GetBATBlockAndIndex(236 * 128).Block.HasFreeSectors);
                Assert.Fail("Should only be 236 BATs");
            }
            catch (ArgumentOutOfRangeException)
            {
            }
            assertBATCount(fs, 236, 1);

            // Ask for another, will get our 2nd XBAT
            free = fs.GetFreeBlock();
            ClassicAssert.AreEqual(false, fs.GetBATBlockAndIndex(236 * 128 - 1).Block.HasFreeSectors);
            ClassicAssert.AreEqual(true, fs.GetBATBlockAndIndex(237 * 128 - 1).Block.HasFreeSectors);
            try
            {
                ClassicAssert.AreEqual(false, fs.GetBATBlockAndIndex(237 * 128).Block.HasFreeSectors);
                Assert.Fail("Should only be 237 BATs");
            }
            // catch (IndexOutOfRangeException) { }
            catch (ArgumentOutOfRangeException)
            {
            }

            // Check the counts now
            assertBATCount(fs, 237, 2);

            // Check the header
            header = WriteOutAndReadHeader(fs);


            // Now, write it out, and read it back in again fully
            fs = WriteOutAndReadBack(fs);

            
            // Check that it is seen correctly
            assertBATCount(fs, 237, 2);

            ClassicAssert.AreEqual(false, fs.GetBATBlockAndIndex(236 * 128 - 1).Block.HasFreeSectors);
            ClassicAssert.AreEqual(true, fs.GetBATBlockAndIndex(237 * 128 - 1).Block.HasFreeSectors);
            try
            {
                ClassicAssert.AreEqual(false, fs.GetBATBlockAndIndex(237 * 128).Block.HasFreeSectors);
                Assert.Fail("Should only be 237 BATs");
            }
            catch (ArgumentOutOfRangeException) { }

            fs.Close();
        }

        /**
        * Test that we can correctly get the list of directory
        *  entries, and the details on the files in them
        */
        [Test]
        public void TestListEntries()
        {
            foreach (NPOIFSFileSystem fs in get512and4kFileAndInput())
            {
                DirectoryEntry root = fs.Root;
                ClassicAssert.AreEqual(5, root.EntryCount);

                // Check by the names
                Entry thumbnail = root.GetEntry("Thumbnail");
                Entry dsi = root.GetEntry("\u0005DocumentSummaryInformation");
                Entry si = root.GetEntry("\u0005SummaryInformation");
                Entry image = root.GetEntry("Image");
                Entry tags = root.GetEntry("Tags");

                ClassicAssert.AreEqual(false, thumbnail.IsDirectoryEntry);
                ClassicAssert.AreEqual(false, dsi.IsDirectoryEntry);
                ClassicAssert.AreEqual(false, si.IsDirectoryEntry);
                ClassicAssert.AreEqual(true, image.IsDirectoryEntry);
                ClassicAssert.AreEqual(false, tags.IsDirectoryEntry);

                // Check via the iterator
                IEnumerator<Entry> it = root.Entries;
                it.MoveNext();
                ClassicAssert.AreEqual(thumbnail.Name, it.Current.Name);
                it.MoveNext();
                ClassicAssert.AreEqual(dsi.Name, it.Current.Name);
                it.MoveNext();
                ClassicAssert.AreEqual(si.Name, it.Current.Name);
                it.MoveNext();
                ClassicAssert.AreEqual(image.Name, it.Current.Name);
                it.MoveNext();
                ClassicAssert.AreEqual(tags.Name, it.Current.Name);

                // Look inside another
                DirectoryEntry imageD = (DirectoryEntry)image;
                ClassicAssert.AreEqual(7, imageD.EntryCount);

                fs.Close();
            }
        }

        [Test]
        public void TestGetDocumentEntry()
        {
            foreach (NPOIFSFileSystem fs in get512and4kFileAndInput())
            {
                DirectoryEntry root = fs.Root;
                Entry si = root.GetEntry("\x0005SummaryInformation");

                ClassicAssert.AreEqual(true, si.IsDocumentEntry);
                DocumentNode doc = (DocumentNode)si;

                // Check we can read it
                assertContentsMatches(null, doc);

                // Now try to build the property set
                //  ByteBuffer temp = inp.GetCurrentBuffer();
                DocumentInputStream inp = new NDocumentInputStream(doc);

                PropertySet ps = PropertySetFactory.Create(inp);
                SummaryInformation inf = (SummaryInformation)ps;

                // Check some bits in it
                ClassicAssert.AreEqual(null, inf.ApplicationName);
                ClassicAssert.AreEqual(null, inf.Author);
                ClassicAssert.AreEqual(null, inf.Subject);
                ClassicAssert.AreEqual(131333, inf.OSVersion);

                // Finish
                inp.Close();

                // Try the other summary information
                si = root.GetEntry("\u0005DocumentSummaryInformation");
                ClassicAssert.AreEqual(true, si.IsDocumentEntry);
                doc = (DocumentNode)si;
                assertContentsMatches(null, doc);

                inp = new NDocumentInputStream(doc);
                ps = PropertySetFactory.Create(inp);
                DocumentSummaryInformation dinf = (DocumentSummaryInformation)ps;
                ClassicAssert.AreEqual(131333, dinf.OSVersion);

                fs.Close();
            }
        }

        /**
            * Read a file, write it and read it again.
            * Then, alter+add some streams, write and read
        */
        [Test]
        public void ReadWriteRead()
        {
            SummaryInformation sinf = null;
            DocumentSummaryInformation dinf = null;
            DirectoryEntry root = null, testDir = null;
            NPOIFSFileSystem[] testFS = get512and4kFileAndInput();
            for (int i=0;i<testFS.Length;i++)
            {
                NPOIFSFileSystem fs = testFS[i];
                // Check we can find the entries we expect
                root = fs.Root;
                ClassicAssert.AreEqual(5, root.EntryCount);
                
                Assert.That(root.EntryNames, new ContainsConstraint("Thumbnail"));
                Assert.That(root.EntryNames, new ContainsConstraint("Image"));
                Assert.That(root.EntryNames, new ContainsConstraint("Tags"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005DocumentSummaryInformation"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005SummaryInformation"));


                // Write out1, re-load
                fs = WriteOutAndReadBack(fs);

                // Check they're still there
                root = fs.Root;
                ClassicAssert.AreEqual(5, root.EntryCount);
                Assert.That(root.EntryNames, new ContainsConstraint("Thumbnail"));
                Assert.That(root.EntryNames, new ContainsConstraint("Image"));
                Assert.That(root.EntryNames, new ContainsConstraint("Tags"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005DocumentSummaryInformation"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005SummaryInformation"));


                // Check the contents of them - parse the summary block and check
                sinf = (SummaryInformation)PropertySetFactory.Create(new NDocumentInputStream(
                        (DocumentEntry)root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME)));
                ClassicAssert.AreEqual(131333, sinf.OSVersion);

                dinf = (DocumentSummaryInformation)PropertySetFactory.Create(new NDocumentInputStream(
                        (DocumentEntry)root.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME)));
                ClassicAssert.AreEqual(131333, dinf.OSVersion);


                // Add a test mini stream
                testDir = root.CreateDirectory("Testing 123");
                testDir.CreateDirectory("Testing 456");
                testDir.CreateDirectory("Testing 789");
                byte[] mini = new byte[] { 42, 0, 1, 2, 3, 4, 42 };
                testDir.CreateDocument("Mini", new MemoryStream(mini));


                // Write out1, re-load
                fs = WriteOutAndReadBack(fs);
                root = fs.Root;
                testDir = (DirectoryEntry)root.GetEntry("Testing 123");
                ClassicAssert.AreEqual(6, root.EntryCount);
                Assert.That(root.EntryNames, new ContainsConstraint("Thumbnail"));
                Assert.That(root.EntryNames, new ContainsConstraint("Image"));
                Assert.That(root.EntryNames, new ContainsConstraint("Tags"));
                Assert.That(root.EntryNames, new ContainsConstraint("Testing 123"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005DocumentSummaryInformation"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005SummaryInformation"));


                // Check old and new are there
                sinf = (SummaryInformation)PropertySetFactory.Create(new NDocumentInputStream(
                        (DocumentEntry)root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME)));
                ClassicAssert.AreEqual(131333, sinf.OSVersion);

                dinf = (DocumentSummaryInformation)PropertySetFactory.Create(new NDocumentInputStream(
                        (DocumentEntry)root.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME)));
                ClassicAssert.AreEqual(131333, dinf.OSVersion);

                assertContentsMatches(mini, (DocumentEntry)testDir.GetEntry("Mini"));


                // Write out and read once more, just to be sure
                fs = WriteOutAndReadBack(fs);

                root = fs.Root;
                testDir = (DirectoryEntry)root.GetEntry("Testing 123");
                ClassicAssert.AreEqual(6, root.EntryCount);
                Assert.That(root.EntryNames, new ContainsConstraint("Thumbnail"));
                Assert.That(root.EntryNames, new ContainsConstraint("Image"));
                Assert.That(root.EntryNames, new ContainsConstraint("Tags"));
                Assert.That(root.EntryNames, new ContainsConstraint("Testing 123"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005DocumentSummaryInformation"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005SummaryInformation"));

                sinf = (SummaryInformation)PropertySetFactory.Create(new NDocumentInputStream(
                        (DocumentEntry)root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME)));
                ClassicAssert.AreEqual(131333, sinf.OSVersion);

                dinf = (DocumentSummaryInformation)PropertySetFactory.Create(new NDocumentInputStream(
                        (DocumentEntry)root.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME)));
                ClassicAssert.AreEqual(131333, dinf.OSVersion);
                assertContentsMatches(mini, (DocumentEntry)testDir.GetEntry("Mini"));

                byte[] main4096 = new byte[4096];
                main4096[0] = unchecked((byte)-10);
                main4096[4095] = unchecked((byte)-11);
                testDir.CreateDocument("Normal4096", new MemoryStream(main4096));

                root.GetEntry("Tags").Delete();


                // Write out1, re-load
                fs = WriteOutAndReadBack(fs);

                // Check it's all there
                root = fs.Root;
                testDir = (DirectoryEntry)root.GetEntry("Testing 123");
                ClassicAssert.AreEqual(5, root.EntryCount);
                Assert.That(root.EntryNames, new ContainsConstraint("Thumbnail"));
                Assert.That(root.EntryNames, new ContainsConstraint("Image"));
                Assert.That(root.EntryNames, new ContainsConstraint("Testing 123"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005DocumentSummaryInformation"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005SummaryInformation"));


                // Check old and new are there
                sinf = (SummaryInformation)PropertySetFactory.Create(new NDocumentInputStream(
                        (DocumentEntry)root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME)));
                ClassicAssert.AreEqual(131333, sinf.OSVersion);

                dinf = (DocumentSummaryInformation)PropertySetFactory.Create(new NDocumentInputStream(
                        (DocumentEntry)root.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME)));
                ClassicAssert.AreEqual(131333, dinf.OSVersion);

                assertContentsMatches(mini, (DocumentEntry)testDir.GetEntry("Mini"));
                assertContentsMatches(main4096, (DocumentEntry)testDir.GetEntry("Normal4096"));


                // Delete a directory, and add one more
                testDir.GetEntry("Testing 456").Delete();
                testDir.CreateDirectory("Testing ABC");


                // Save
                fs = WriteOutAndReadBack(fs);

                // Check
                root = fs.Root;
                testDir = (DirectoryEntry)root.GetEntry("Testing 123");

                ClassicAssert.AreEqual(5, root.EntryCount);
                Assert.That(root.EntryNames, new ContainsConstraint("Thumbnail"));
                Assert.That(root.EntryNames, new ContainsConstraint("Image"));
                Assert.That(root.EntryNames, new ContainsConstraint("Testing 123"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005DocumentSummaryInformation"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005SummaryInformation"));

                ClassicAssert.AreEqual(4, testDir.EntryCount);
                Assert.That(testDir.EntryNames, new ContainsConstraint("Mini"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Normal4096"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Testing 789"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Testing ABC"));


                // Add another mini stream
                byte[] mini2 = new byte[] { unchecked((byte)-42), 0, unchecked((byte)-1), 
                    unchecked((byte)-2), unchecked((byte)-3), unchecked((byte)-4), unchecked((byte)-42) };
                testDir.CreateDocument("Mini2", new MemoryStream(mini2));

                // Save, load, check
                fs = WriteOutAndReadBack(fs);

                root = fs.Root;
                testDir = (DirectoryEntry)root.GetEntry("Testing 123");

                ClassicAssert.AreEqual(5, root.EntryCount);
                Assert.That(root.EntryNames, new ContainsConstraint("Thumbnail"));
                Assert.That(root.EntryNames, new ContainsConstraint("Image"));
                Assert.That(root.EntryNames, new ContainsConstraint("Testing 123"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005DocumentSummaryInformation"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005SummaryInformation"));

                ClassicAssert.AreEqual(5, testDir.EntryCount);
                Assert.That(testDir.EntryNames, new ContainsConstraint("Mini"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Mini2"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Normal4096"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Testing 789"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Testing ABC"));

                assertContentsMatches(mini, (DocumentEntry)testDir.GetEntry("Mini"));
                assertContentsMatches(mini2, (DocumentEntry)testDir.GetEntry("Mini2"));
                assertContentsMatches(main4096, (DocumentEntry)testDir.GetEntry("Normal4096"));


                // Delete a mini stream, add one more
                testDir.GetEntry("Mini").Delete();

                byte[] mini3 = new byte[] { 42, 0, 42, 0, 42, 0, 42 };
                testDir.CreateDocument("Mini3", new MemoryStream(mini3));


                // Save, load, check
                fs = WriteOutAndReadBack(fs);

                root = fs.Root;
                testDir = (DirectoryEntry)root.GetEntry("Testing 123");

                ClassicAssert.AreEqual(5, root.EntryCount);
                Assert.That(root.EntryNames, new ContainsConstraint("Thumbnail"));
                Assert.That(root.EntryNames, new ContainsConstraint("Image"));
                Assert.That(root.EntryNames, new ContainsConstraint("Testing 123"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005DocumentSummaryInformation"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005SummaryInformation"));

                ClassicAssert.AreEqual(5, testDir.EntryCount);
                Assert.That(testDir.EntryNames, new ContainsConstraint("Mini2"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Mini3"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Normal4096"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Testing 789"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Testing ABC"));

                assertContentsMatches(mini2, (DocumentEntry)testDir.GetEntry("Mini2"));
                assertContentsMatches(mini3, (DocumentEntry)testDir.GetEntry("Mini3"));
                assertContentsMatches(main4096, (DocumentEntry)testDir.GetEntry("Normal4096"));

                // Change some existing streams
                NPOIFSDocument mini2Doc = new NPOIFSDocument((DocumentNode)testDir.GetEntry("Mini2"));
                mini2Doc.ReplaceContents(new MemoryStream(mini));

                byte[] main4106 = new byte[4106];
                main4106[0] = 41;
                main4106[4105] = 42;
                NPOIFSDocument mainDoc = new NPOIFSDocument((DocumentNode)testDir.GetEntry("Normal4096"));
                mainDoc.ReplaceContents(new MemoryStream(main4106));


                // Re-check
                fs = WriteOutAndReadBack(fs);

                root = fs.Root;
                testDir = (DirectoryEntry)root.GetEntry("Testing 123");

                ClassicAssert.AreEqual(5, root.EntryCount);
                Assert.That(root.EntryNames, new ContainsConstraint("Thumbnail"));
                Assert.That(root.EntryNames, new ContainsConstraint("Image"));
                Assert.That(root.EntryNames, new ContainsConstraint("Testing 123"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005DocumentSummaryInformation"));
                Assert.That(root.EntryNames, new ContainsConstraint("\u0005SummaryInformation"));

                ClassicAssert.AreEqual(5, testDir.EntryCount);
                Assert.That(testDir.EntryNames, new ContainsConstraint("Mini2"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Mini3"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Normal4096"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Testing 789"));
                Assert.That(testDir.EntryNames, new ContainsConstraint("Testing ABC"));

                assertContentsMatches(mini, (DocumentEntry)testDir.GetEntry("Mini2"));
                assertContentsMatches(mini3, (DocumentEntry)testDir.GetEntry("Mini3"));
                assertContentsMatches(main4106, (DocumentEntry)testDir.GetEntry("Normal4096"));
           


                fs.Close();
            }
        }

        /**
             * Create a new file, write it and read it again
             * Then, add some streams, write and read
        */
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void CreateWriteRead()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem();
            DocumentEntry miniDoc;
            DocumentEntry normDoc;

            // Initially has Properties + BAT but not SBAT
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(0));
            ClassicAssert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(1));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(2));

            // Check that the SBAT is empty
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.Root.Property.StartBlock);

            // Check that properties table was given block 0
            ClassicAssert.AreEqual(0, fs.PropertyTable.StartBlock);

            // Write and read it
            fs = WriteOutAndReadBack(fs);

            // No change, SBAT remains empty 
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(0));
            ClassicAssert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(1));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(2));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(3));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.Root.Property.StartBlock);
            ClassicAssert.AreEqual(0, fs.PropertyTable.StartBlock);

            // Check the same but with saving to a file
            fs = new NPOIFSFileSystem();
            fs = WriteOutFileAndReadBack(fs);

            // Same, no change, SBAT remains empty 
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(0));
            ClassicAssert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(1));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(2));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(3));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.Root.Property.StartBlock);
            ClassicAssert.AreEqual(0, fs.PropertyTable.StartBlock);


            // Put everything within a new directory
            DirectoryEntry testDir = fs.CreateDirectory("Test Directory");

            // Add a new Normal Stream (Normal Streams minimum 4096 bytes)
            byte[] main4096 = new byte[4096];
            main4096[0] = unchecked((byte)-10);
            main4096[4095] = unchecked((byte)-11);
            testDir.CreateDocument("Normal4096", new MemoryStream(main4096));

            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(0));
            ClassicAssert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(1));
            ClassicAssert.AreEqual(3, fs.GetNextBlock(2));
            ClassicAssert.AreEqual(4, fs.GetNextBlock(3));
            ClassicAssert.AreEqual(5, fs.GetNextBlock(4));
            ClassicAssert.AreEqual(6, fs.GetNextBlock(5));
            ClassicAssert.AreEqual(7, fs.GetNextBlock(6));
            ClassicAssert.AreEqual(8, fs.GetNextBlock(7));
            ClassicAssert.AreEqual(9, fs.GetNextBlock(8));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(9));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(10));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(11));
            // SBAT still unused
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.Root.Property.StartBlock);


            // Add a bigger Normal Stream
            byte[] main5124 = new byte[5124];
            main5124[0] = unchecked((byte)-22);
            main5124[5123] = unchecked((byte)-33);
            testDir.CreateDocument("Normal5124", new MemoryStream(main5124));

            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(0));
            ClassicAssert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(1));
            ClassicAssert.AreEqual(3, fs.GetNextBlock(2));
            ClassicAssert.AreEqual(4, fs.GetNextBlock(3));
            ClassicAssert.AreEqual(5, fs.GetNextBlock(4));
            ClassicAssert.AreEqual(6, fs.GetNextBlock(5));
            ClassicAssert.AreEqual(7, fs.GetNextBlock(6));
            ClassicAssert.AreEqual(8, fs.GetNextBlock(7));
            ClassicAssert.AreEqual(9, fs.GetNextBlock(8));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(9));

            ClassicAssert.AreEqual(11, fs.GetNextBlock(10));
            ClassicAssert.AreEqual(12, fs.GetNextBlock(11));
            ClassicAssert.AreEqual(13, fs.GetNextBlock(12));
            ClassicAssert.AreEqual(14, fs.GetNextBlock(13));
            ClassicAssert.AreEqual(15, fs.GetNextBlock(14));
            ClassicAssert.AreEqual(16, fs.GetNextBlock(15));
            ClassicAssert.AreEqual(17, fs.GetNextBlock(16));
            ClassicAssert.AreEqual(18, fs.GetNextBlock(17));
            ClassicAssert.AreEqual(19, fs.GetNextBlock(18));
            ClassicAssert.AreEqual(20, fs.GetNextBlock(19));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(20));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(21));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(22));

            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.Root.Property.StartBlock);


            // Now Add a mini stream
            byte[] mini = new byte[] { 42, 0, 1, 2, 3, 4, 42 };
            testDir.CreateDocument("Mini", new MemoryStream(mini));

            // Mini stream will Get one block for fat + one block for data
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(0));
            ClassicAssert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(1));
            ClassicAssert.AreEqual(3, fs.GetNextBlock(2));
            ClassicAssert.AreEqual(4, fs.GetNextBlock(3));
            ClassicAssert.AreEqual(5, fs.GetNextBlock(4));
            ClassicAssert.AreEqual(6, fs.GetNextBlock(5));
            ClassicAssert.AreEqual(7, fs.GetNextBlock(6));
            ClassicAssert.AreEqual(8, fs.GetNextBlock(7));
            ClassicAssert.AreEqual(9, fs.GetNextBlock(8));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(9));

            ClassicAssert.AreEqual(11, fs.GetNextBlock(10));
            ClassicAssert.AreEqual(12, fs.GetNextBlock(11));
            ClassicAssert.AreEqual(13, fs.GetNextBlock(12));
            ClassicAssert.AreEqual(14, fs.GetNextBlock(13));
            ClassicAssert.AreEqual(15, fs.GetNextBlock(14));
            ClassicAssert.AreEqual(16, fs.GetNextBlock(15));
            ClassicAssert.AreEqual(17, fs.GetNextBlock(16));
            ClassicAssert.AreEqual(18, fs.GetNextBlock(17));
            ClassicAssert.AreEqual(19, fs.GetNextBlock(18));
            ClassicAssert.AreEqual(20, fs.GetNextBlock(19));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(20));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(21));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(22));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(23));

            // Check the mini stream location was set
            // (21 is mini fat, 22 is first mini stream block)
            ClassicAssert.AreEqual(22, fs.Root.Property.StartBlock);

            // Write and read back
            fs = WriteOutAndReadBack(fs);
            HeaderBlock header = WriteOutAndReadHeader(fs);

            // Check the header has the right points in it
            ClassicAssert.AreEqual(1, header.BATCount);
            ClassicAssert.AreEqual(1, header.BATArray[0]);
            ClassicAssert.AreEqual(0, header.PropertyStart);
            ClassicAssert.AreEqual(1, header.SBATCount);
            ClassicAssert.AreEqual(21, header.SBATStart);
            ClassicAssert.AreEqual(22, fs.PropertyTable.Root.StartBlock);

            // Block use should be almost the same, except the properties
            //  stream will have grown out to cover 2 blocks
            // Check the block use is all unChanged
            // Check it's all unChanged
            ClassicAssert.AreEqual(23, fs.GetNextBlock(0));// Properties now extends over 2 blocks
            ClassicAssert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(1));

            ClassicAssert.AreEqual(3, fs.GetNextBlock(2));
            ClassicAssert.AreEqual(4, fs.GetNextBlock(3));
            ClassicAssert.AreEqual(5, fs.GetNextBlock(4));
            ClassicAssert.AreEqual(6, fs.GetNextBlock(5));
            ClassicAssert.AreEqual(7, fs.GetNextBlock(6));
            ClassicAssert.AreEqual(8, fs.GetNextBlock(7));
            ClassicAssert.AreEqual(9, fs.GetNextBlock(8));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(9));// End of normal4096

            ClassicAssert.AreEqual(11, fs.GetNextBlock(10));
            ClassicAssert.AreEqual(12, fs.GetNextBlock(11));
            ClassicAssert.AreEqual(13, fs.GetNextBlock(12));
            ClassicAssert.AreEqual(14, fs.GetNextBlock(13));
            ClassicAssert.AreEqual(15, fs.GetNextBlock(14));
            ClassicAssert.AreEqual(16, fs.GetNextBlock(15));
            ClassicAssert.AreEqual(17, fs.GetNextBlock(16));
            ClassicAssert.AreEqual(18, fs.GetNextBlock(17));
            ClassicAssert.AreEqual(19, fs.GetNextBlock(18));
            ClassicAssert.AreEqual(20, fs.GetNextBlock(19));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(20)); // End of normal5124 

            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(21)); // Mini Stream FAT
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(22)); // Mini Stream data
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(23)); // Properties #2
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(24));

            // Check some data
            ClassicAssert.AreEqual(1, fs.Root.EntryCount);
            testDir = (DirectoryEntry)fs.Root.GetEntry("Test Directory");
            ClassicAssert.AreEqual(3, testDir.EntryCount);

            miniDoc = (DocumentEntry)testDir.GetEntry("Mini");
            assertContentsMatches(mini, miniDoc);

            normDoc = (DocumentEntry)testDir.GetEntry("Normal4096");
            assertContentsMatches(main4096, normDoc);

            normDoc = (DocumentEntry)testDir.GetEntry("Normal5124");
            assertContentsMatches(main5124, normDoc);

            // Delete a couple of streams
            miniDoc.Delete();
            normDoc.Delete();


            // Check - will have un-used sectors now
            fs = WriteOutAndReadBack(fs);

            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(0));// Props back in 1 block
            ClassicAssert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(1));

            ClassicAssert.AreEqual(3, fs.GetNextBlock(2)); 
            ClassicAssert.AreEqual(4, fs.GetNextBlock(3));
            ClassicAssert.AreEqual(5, fs.GetNextBlock(4));
            ClassicAssert.AreEqual(6, fs.GetNextBlock(5));
            ClassicAssert.AreEqual(7, fs.GetNextBlock(6));
            ClassicAssert.AreEqual(8, fs.GetNextBlock(7));
            ClassicAssert.AreEqual(9, fs.GetNextBlock(8));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(9));  // End of normal4096

            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(10));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(11));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(12));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(13));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(14));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(15));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(16));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(17));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(18));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(19));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(20));

            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(21)); // Mini Stream FAT
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(22)); // Mini Stream data
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(23)); // Properties gone
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(24));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(25));
      

            // All done
            fs.Close();
        }

        [Test]
        public void AddBeforeWrite()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem();
            DocumentEntry miniDoc;
            DocumentEntry normDoc;
            HeaderBlock hdr;

            // Initially has Properties + BAT but nothing else
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(0));
            ClassicAssert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(1));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(2));

            hdr = WriteOutAndReadHeader(fs);
            // No mini stream, and no xbats
            // Will have fat then properties stream
            ClassicAssert.AreEqual(1, hdr.BATCount);
            ClassicAssert.AreEqual(1, hdr.BATArray[0]);
            ClassicAssert.AreEqual(0, hdr.PropertyStart);
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, hdr.SBATStart);
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, hdr.XBATIndex);
            ClassicAssert.AreEqual(POIFSConstants.SMALLER_BIG_BLOCK_SIZE * 3, fs.Size);


            // Get a clean filesystem to start with
            fs = new NPOIFSFileSystem();

            // Put our test files in a non-standard place
            DirectoryEntry parentDir = fs.CreateDirectory("Parent Directory");
            DirectoryEntry testDir = parentDir.CreateDirectory("Test Directory");


            // Add to the mini stream
            byte[] mini = new byte[] { 42, 0, 1, 2, 3, 4, 42 };
            testDir.CreateDocument("Mini", new MemoryStream(mini));

            // Add to the main stream
            byte[] main4096 = new byte[4096];
            main4096[0] = unchecked((byte)-10);
            main4096[4095] = unchecked((byte)-11);
            testDir.CreateDocument("Normal4096", new MemoryStream(main4096));


            // Check the mini stream was Added, then the main stream
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(0));
            ClassicAssert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(1));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(2)); // Mini Fat
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(3)); // Mini Stream
            ClassicAssert.AreEqual(5, fs.GetNextBlock(4)); // Main Stream
            ClassicAssert.AreEqual(6, fs.GetNextBlock(5));
            ClassicAssert.AreEqual(7, fs.GetNextBlock(6));
            ClassicAssert.AreEqual(8, fs.GetNextBlock(7));
            ClassicAssert.AreEqual(9, fs.GetNextBlock(8));
            ClassicAssert.AreEqual(10, fs.GetNextBlock(9));
            ClassicAssert.AreEqual(11, fs.GetNextBlock(10));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(11));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(12));
            ClassicAssert.AreEqual(POIFSConstants.SMALLER_BIG_BLOCK_SIZE * 13, fs.Size);

            // Check that we can read the right data pre-write
            miniDoc = (DocumentEntry)testDir.GetEntry("Mini");
            assertContentsMatches(mini, miniDoc);

            normDoc = (DocumentEntry)testDir.GetEntry("Normal4096");
            assertContentsMatches(main4096, normDoc);

            // Write, Read, check
            hdr = WriteOutAndReadHeader(fs);
            fs = WriteOutAndReadBack(fs);

            // Check the header details - will have the sbat near the start,
            //  then the properties at the end
            ClassicAssert.AreEqual(1, hdr.BATCount);
            ClassicAssert.AreEqual(1, hdr.BATArray[0]);
            ClassicAssert.AreEqual(2, hdr.SBATStart);
            ClassicAssert.AreEqual(0, hdr.PropertyStart);
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, hdr.XBATIndex);

            // Check the block allocation is unChanged, other than
            //  the properties stream going in at the end
            ClassicAssert.AreEqual(12, fs.GetNextBlock(0));
            ClassicAssert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(1)); // Properties
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(2));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(3));
            ClassicAssert.AreEqual(5, fs.GetNextBlock(4));
            ClassicAssert.AreEqual(6, fs.GetNextBlock(5));
            ClassicAssert.AreEqual(7, fs.GetNextBlock(6));
            ClassicAssert.AreEqual(8, fs.GetNextBlock(7));
            ClassicAssert.AreEqual(9, fs.GetNextBlock(8));
            ClassicAssert.AreEqual(10, fs.GetNextBlock(9));
            ClassicAssert.AreEqual(11, fs.GetNextBlock(10));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(11));
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(12));
            ClassicAssert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(13));
            ClassicAssert.AreEqual(POIFSConstants.SMALLER_BIG_BLOCK_SIZE * 14, fs.Size);
       
            // Check the data
            DirectoryEntry fsRoot = fs.Root;
            ClassicAssert.AreEqual(1, fsRoot.EntryCount);

            parentDir = (DirectoryEntry)fsRoot.GetEntry("Parent Directory");
            ClassicAssert.AreEqual(1, parentDir.EntryCount);

            testDir = (DirectoryEntry)parentDir.GetEntry("Test Directory");
            ClassicAssert.AreEqual(2, testDir.EntryCount);

            miniDoc = (DocumentEntry)testDir.GetEntry("Mini");
            assertContentsMatches(mini, miniDoc);


            normDoc = (DocumentEntry)testDir.GetEntry("Normal4096");
            assertContentsMatches(main4096, normDoc);

            byte[] mini2 = new byte[] { unchecked((byte)-42), 0, unchecked((byte)-1),
                unchecked((byte)-2), unchecked((byte)-3), unchecked((byte)-4), unchecked((byte)-42) };
            testDir.CreateDocument("Mini2", new MemoryStream(mini2));

            // Add to the main stream
            byte[] main4106 = new byte[4106];
            main4106[0] = 41;
            main4106[4105] = 42;
            testDir.CreateDocument("Normal4106", new MemoryStream(main4106));


            // Recheck the data in all 4 streams
            fs = WriteOutAndReadBack(fs);

            fsRoot = fs.Root;
            ClassicAssert.AreEqual(1, fsRoot.EntryCount);

            parentDir = (DirectoryEntry)fsRoot.GetEntry("Parent Directory");
            ClassicAssert.AreEqual(1, parentDir.EntryCount);

            testDir = (DirectoryEntry)parentDir.GetEntry("Test Directory");
            ClassicAssert.AreEqual(4, testDir.EntryCount);

            miniDoc = (DocumentEntry)testDir.GetEntry("Mini");
            assertContentsMatches(mini, miniDoc);

            miniDoc = (DocumentEntry)testDir.GetEntry("Mini2");
            assertContentsMatches(mini2, miniDoc);

            normDoc = (DocumentEntry)testDir.GetEntry("Normal4106");
            assertContentsMatches(main4106, normDoc);

            // All done
            fs.Close();
        }
        [Test]
        public void ReadZeroLengthEntries()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.GetFile("only-zero-byte-streams.ole2"));
            DirectoryNode testDir = fs.Root;
            ClassicAssert.AreEqual(3, testDir.EntryCount);
            DocumentEntry entry;

            entry = (DocumentEntry)testDir.GetEntry("test-zero-1");
            ClassicAssert.IsNotNull(entry);
            ClassicAssert.AreEqual(0, entry.Size);

            entry = (DocumentEntry)testDir.GetEntry("test-zero-2");
            ClassicAssert.IsNotNull(entry);
            ClassicAssert.AreEqual(0, entry.Size);

            entry = (DocumentEntry)testDir.GetEntry("test-zero-3");
            ClassicAssert.IsNotNull(entry);
            ClassicAssert.AreEqual(0, entry.Size);

            // Check properties, all have zero length, no blocks
            NPropertyTable props = fs.PropertyTable;
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, props.Root.StartBlock);
            foreach (NPOI.POIFS.Properties.Property prop in props.Root)
            {
                ClassicAssert.AreEqual("test-zero-", prop.Name.Substring(0, 10));
                ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, prop.StartBlock);
            }

            // All done
            fs.Close();
        }

        [Test]
        public void WriteZeroLengthEntries()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem();
            DirectoryNode testDir = fs.Root;
            DocumentEntry miniDoc;
            DocumentEntry normDoc;
            DocumentEntry emptyDoc;

            byte[] mini2;
            // Add mini and normal sized entries to start
            unchecked
            {
                mini2 = new byte[] { (byte)-42, 0, (byte)-1, (byte)-2, (byte)-3, (byte)-4, (byte)-42 };
            }
            
            testDir.CreateDocument("Mini2", new ByteArrayInputStream(mini2));

            // Add to the main stream
            byte[] main4106 = new byte[4106];
            main4106[0] = 41;
            main4106[4105] = 42;
            testDir.CreateDocument("Normal4106", new ByteArrayInputStream(main4106));

            // Now add some empty ones
            byte[] empty = new byte[0];
            testDir.CreateDocument("empty-1", new ByteArrayInputStream(empty));
            testDir.CreateDocument("empty-2", new ByteArrayInputStream(empty));
            testDir.CreateDocument("empty-3", new ByteArrayInputStream(empty));

            // Check
            miniDoc = (DocumentEntry)testDir.GetEntry("Mini2");
            assertContentsMatches(mini2, miniDoc);

            normDoc = (DocumentEntry)testDir.GetEntry("Normal4106");
            assertContentsMatches(main4106, normDoc);

            emptyDoc = (DocumentEntry)testDir.GetEntry("empty-1");
            assertContentsMatches(empty, emptyDoc);

            emptyDoc = (DocumentEntry)testDir.GetEntry("empty-2");
            assertContentsMatches(empty, emptyDoc);

            emptyDoc = (DocumentEntry)testDir.GetEntry("empty-3");
            assertContentsMatches(empty, emptyDoc);

            // Look at the properties entry, and check the empty ones
            //  have zero size and no start block
            NPropertyTable props = fs.PropertyTable;
            IEnumerator<Property> propsIt = props.Root.Children;

            propsIt.MoveNext();
            Property prop = propsIt.Current;
            ClassicAssert.AreEqual("Mini2", prop.Name);
            ClassicAssert.AreEqual(0, prop.StartBlock);
            ClassicAssert.AreEqual(7, prop.Size);

            propsIt.MoveNext();
            prop = propsIt.Current;
            ClassicAssert.AreEqual("Normal4106", prop.Name);
            ClassicAssert.AreEqual(4, prop.StartBlock); // BAT, Props, SBAT, MIni
            ClassicAssert.AreEqual(4106, prop.Size);

            propsIt.MoveNext();
            prop = propsIt.Current;
            ClassicAssert.AreEqual("empty-1", prop.Name);
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, prop.StartBlock);
            ClassicAssert.AreEqual(0, prop.Size);

            propsIt.MoveNext();
            prop = propsIt.Current;
            ClassicAssert.AreEqual("empty-2", prop.Name);
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, prop.StartBlock);
            ClassicAssert.AreEqual(0, prop.Size);

            propsIt.MoveNext();
            prop = propsIt.Current;
            ClassicAssert.AreEqual("empty-3", prop.Name);
            ClassicAssert.AreEqual(POIFSConstants.END_OF_CHAIN, prop.StartBlock);
            ClassicAssert.AreEqual(0, prop.Size);

            // Save and re-check
            fs = WriteOutAndReadBack(fs);
            testDir = fs.Root;

            miniDoc = (DocumentEntry)testDir.GetEntry("Mini2");
            assertContentsMatches(mini2, miniDoc);

            normDoc = (DocumentEntry)testDir.GetEntry("Normal4106");
            assertContentsMatches(main4106, normDoc);

            emptyDoc = (DocumentEntry)testDir.GetEntry("empty-1");
            assertContentsMatches(empty, emptyDoc);

            emptyDoc = (DocumentEntry)testDir.GetEntry("empty-2");
            assertContentsMatches(empty, emptyDoc);

            emptyDoc = (DocumentEntry)testDir.GetEntry("empty-3");
            assertContentsMatches(empty, emptyDoc);

            // Check that a mini-stream was assigned, with one block used
            ClassicAssert.AreEqual(3, testDir.Property.StartBlock);
            ClassicAssert.AreEqual(64, testDir.Property.Size);
            // All done
            fs.Close();
        }

        /**
        * Test that we can read a file with NPOIFS, create a new NPOIFS instance,
        *  write it out, read it with POIFS, and see the original data
        */
        [Test]
        public void NPOIFSReadCopyWritePOIFSRead()
        {
            FileStream testFile = POIDataSamples.GetSpreadSheetInstance().GetFile("Simple.xls");
            NPOIFSFileSystem src = new NPOIFSFileSystem(testFile);
            byte[] wbDataExp = IOUtils.ToByteArray(src.CreateDocumentInputStream("Workbook"));

            NPOIFSFileSystem nfs = new NPOIFSFileSystem();
            EntryUtils.CopyNodes(src.Root, nfs.Root);
            src.Close();

            MemoryStream bos = new MemoryStream();
            nfs.WriteFileSystem(bos);
            nfs.Close();

            POIFSFileSystem pfs = new POIFSFileSystem(new MemoryStream(bos.ToArray()));
            byte[] wbDataAct = IOUtils.ToByteArray(pfs.CreateDocumentInputStream("Workbook"));

            Assert.That(wbDataExp, new EqualConstraint(wbDataAct));
        }

        /**
        * Ensure that you can recursively delete directories and their
        *  contents
        */
        [Test]
        public void RecursiveDelete() 
        {
            FileStream testFile = POIDataSamples.GetSpreadSheetInstance().GetFile("SimpleMacro.xls");
            NPOIFSFileSystem src = new NPOIFSFileSystem(testFile);

            // Starts out with 5 entries:
            //  _VBA_PROJECT_CUR
            //  SummaryInformation <(0x05)SummaryInformation>
            //  DocumentSummaryInformation <(0x05)DocumentSummaryInformation>
            //  Workbook
            //  CompObj <(0x01)CompObj>
            ClassicAssert.AreEqual(5, _countChildren(src.PropertyTable.Root));
            ClassicAssert.AreEqual(5, src.Root.EntryCount);

            // Grab the VBA project root
            DirectoryEntry vbaProj = (DirectoryEntry)src.Root.GetEntry("_VBA_PROJECT_CUR");
            ClassicAssert.AreEqual(3, vbaProj.EntryCount);
            // Can't delete yet, has stuff
            ClassicAssert.AreEqual(false, vbaProj.Delete());
            // Recursively delete
            _recursiveDeletee(vbaProj);

            // Entries gone
            ClassicAssert.AreEqual(4, _countChildren(src.PropertyTable.Root));
            ClassicAssert.AreEqual(4, src.Root.EntryCount);

            // Done
            src.Close();
        }
        private void _recursiveDeletee(Entry entry)
        {
            if (entry.IsDocumentEntry)
            {
                ClassicAssert.AreEqual(true, entry.Delete());
                return;
            }

            DirectoryEntry dir = (DirectoryEntry)entry;
            String[] names = dir.EntryNames.ToArray();
            foreach (String name in names)
            {
                Entry ce = dir.GetEntry(name);
                _recursiveDeletee(ce);
            }
            ClassicAssert.AreEqual(true, dir.Delete());
        }

        private int _countChildren(DirectoryProperty p)
        {
            int count = 0;
            foreach (Property cp in p) { count++; }
            return count;
        }
    }
}