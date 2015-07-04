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

using NUnit.Framework;
using NPOI.POIFS.Storage;
using NPOI.POIFS.Properties;
using NPOI.POIFS.FileSystem;
using NPOI.POIFS.Common;
using NPOI.HPSF;
using NPOI.POIFS.NIO;
using NPOI.Util;
using NPOI.POIFS.EventFileSystem;

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
            Assert.AreEqual(expectedBAT, foundBAT, "Wrong number of BATs");
            Assert.AreEqual(expectedXBAT, foundXBAT, "Wrong number of XBATs with " + expectedBAT + " BATs");
        }

        protected static HeaderBlock WriteOutAndReadHeader(NPOIFSFileSystem fs)
        {
            MemoryStream baos = new MemoryStream();
            fs.WriteFilesystem(baos);

            HeaderBlock header = new HeaderBlock(new MemoryStream(baos.ToArray()));
            return header;
        }

        [Test]
        public void TestBasicOpen()
        {
            NPOIFSFileSystem fsA, fsB;
            fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));
            fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));

            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB })
            {
                Assert.AreEqual(512, fs.GetBigBlockSize());
            }

            fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize4096.zvi"));
            fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize4096.zvi"));

            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB })
            {
                Assert.AreEqual(4096, fs.GetBigBlockSize());
            }
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

                Assert.AreEqual(98, fs.GetNextBlock(97));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(98));

                NPropertyTable props = fs.PropertyTable;
                Assert.AreEqual(90, props.StartBlock);
                Assert.AreEqual(7, props.CountBlocks);

                RootProperty root = props.Root;
                Assert.AreEqual("Root Entry", root.Name);
                Assert.AreEqual(11564, root.Size);
                Assert.AreEqual(0, root.StartBlock);

                NPOI.POIFS.Properties.Property prop;
                IEnumerator<NPOI.POIFS.Properties.Property> pi = root.Children;
                //prop = pi.Current;
                pi.MoveNext();
                prop = pi.Current;
                Assert.AreEqual("Thumbnail", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                Assert.AreEqual("\x0005DocumentSummaryInformation", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                Assert.AreEqual("\x0005SummaryInformation", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                Assert.AreEqual("Image", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                Assert.AreEqual(false, pi.MoveNext());

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
                    Assert.AreEqual(i + 1, miniStore.GetNextBlock(i));

                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, miniStore.GetNextBlock(50));


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

                Assert.AreEqual(1, fs.GetNextBlock(0));
                Assert.AreEqual(2, fs.GetNextBlock(1));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(2));

                NPropertyTable props = fs.PropertyTable;
                Assert.AreEqual(12, props.StartBlock);
                Assert.AreEqual(1, props.CountBlocks);

                RootProperty root = props.Root;
                Assert.AreEqual("Root Entry", root.Name);
                Assert.AreEqual(11564, root.Size);
                Assert.AreEqual(0, root.StartBlock);

                NPOI.POIFS.Properties.Property prop;
                IEnumerator<NPOI.POIFS.Properties.Property> pi = root.Children;

                pi.MoveNext();
                prop = pi.Current;
                Assert.AreEqual("Thumbnail", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                Assert.AreEqual("\x0005DocumentSummaryInformation", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                Assert.AreEqual("\x0005SummaryInformation", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                Assert.AreEqual("Image", prop.Name);

                pi.MoveNext();
                prop = pi.Current;
                Assert.AreEqual("Tags", prop.Name);
                Assert.AreEqual(false, pi.MoveNext());

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
                    Assert.AreEqual(i + 1, miniStore.GetNextBlock(i));

                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, miniStore.GetNextBlock(50));
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
                    Assert.AreEqual(i + 1, fs.GetNextBlock(i));

                // 21 jumps to 89, then ends
                Assert.AreEqual(89, fs.GetNextBlock(21));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(89));

                for (int i = 22; i < 88; i++)
                    Assert.AreEqual(i + 1, fs.GetNextBlock(i));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(88));

                for (int i = 90; i < 96; i++)
                {
                    Assert.AreEqual(i + 1, fs.GetNextBlock(i));
                }
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(96));

                Assert.AreEqual(98, fs.GetNextBlock(97));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(98));

                Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(99));

                //Leon i = 100
                for (int i = 100; i < fs.GetBigBlockSizeDetails().GetBATEntriesPerBlock(); i++)
                    Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(i));
            }

            fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize4096.zvi"));
            fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize4096.zvi"));

            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB })
            {
                Assert.AreEqual(1, fs.GetNextBlock(0));
                Assert.AreEqual(2, fs.GetNextBlock(1));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(2));

                for (int i = 4; i < 11; i++)
                    Assert.AreEqual(i + 1, fs.GetNextBlock(i));

                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(11));
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
                Assert.AreEqual((byte)0x9e, b.Read());
                Assert.AreEqual((byte)0x75, b.Read());
                Assert.AreEqual((byte)0x97, b.Read());
                Assert.AreEqual((byte)0xf6, b.Read());

                b = fs.GetBlockAt(1);
                Assert.AreEqual((byte)0x86, b.Read());
                Assert.AreEqual((byte)0x09, b.Read());
                Assert.AreEqual((byte)0x22, b.Read());
                Assert.AreEqual((byte)0xfb, b.Read());

                b = fs.GetBlockAt(99);
                Assert.AreEqual((byte)0x01, b.Read());
                Assert.AreEqual((byte)0x00, b.Read());
                Assert.AreEqual((byte)0x00, b.Read());
                Assert.AreEqual((byte)0x00, b.Read());
                Assert.AreEqual((byte)0x02, b.Read());
                Assert.AreEqual((byte)0x00, b.Read());
                Assert.AreEqual((byte)0x00, b.Read());
                Assert.AreEqual((byte)0x00, b.Read());
            }

            fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize4096.zvi"));
            fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize4096.zvi"));


            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB })
            {
                ByteBuffer b;

                b = fs.GetBlockAt(0);
                Assert.AreEqual((byte)0x9e, b.Read());
                Assert.AreEqual((byte)0x75, b.Read());
                Assert.AreEqual((byte)0x97, b.Read());
                Assert.AreEqual((byte)0xf6, b.Read());

                b = fs.GetBlockAt(1);
                Assert.AreEqual((byte)0x00, b.Read());
                Assert.AreEqual((byte)0x00, b.Read());
                Assert.AreEqual((byte)0x03, b.Read());
                Assert.AreEqual((byte)0x00, b.Read());

                b = fs.GetBlockAt(14);
                Assert.AreEqual((byte)0x01, b.Read());
                Assert.AreEqual((byte)0x00, b.Read());
                Assert.AreEqual((byte)0x00, b.Read());
                Assert.AreEqual((byte)0x00, b.Read());
                Assert.AreEqual((byte)0x02, b.Read());
                Assert.AreEqual((byte)0x00, b.Read());
                Assert.AreEqual((byte)0x00, b.Read());
                Assert.AreEqual((byte)0x00, b.Read());
            }
        }

        [Test]
        public void TestGetFreeBlockWithSpare()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));

            Assert.AreEqual(true, fs.GetBATBlockAndIndex(0).Block.HasFreeSectors);

            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(100));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(101));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(102));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(103));

            Assert.AreEqual(100, fs.GetFreeBlock());

            Assert.AreEqual(100, fs.GetFreeBlock());

            fs.SetNextBlock(100, POIFSConstants.END_OF_CHAIN);
            Assert.AreEqual(101, fs.GetFreeBlock());

            fs.Close();
        }

        [Test]
        public void TestGetFreeBlockWithNoneSpare()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));
            int free;

            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(99));

            for (int i = 100; i < 128; i++)
                Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(i));

            Assert.AreEqual(true, fs.GetBATBlockAndIndex(0).Block.HasFreeSectors);

            for (int i = 100; i < 128; i++)
                fs.SetNextBlock(i, POIFSConstants.END_OF_CHAIN);

            Assert.AreEqual(false, fs.GetBATBlockAndIndex(0).Block.HasFreeSectors);

            try
            {
                Assert.AreEqual(false, fs.GetBATBlockAndIndex(128).Block.HasFreeSectors);
                Assert.Fail("Should only be one BAT");
            }
            //catch (IndexOutOfRangeException)
            catch (ArgumentOutOfRangeException)
            {
            }
            assertBATCount(fs, 1, 0);

            // Now ask for a free one, will need to extend the file

            Assert.AreEqual(129, fs.GetFreeBlock());

            Assert.AreEqual(false, fs.GetBATBlockAndIndex(0).Block.HasFreeSectors);
            Assert.AreEqual(true, fs.GetBATBlockAndIndex(128).Block.HasFreeSectors);
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(128));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(129));

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

            Assert.AreEqual(false, fs.GetBATBlockAndIndex(109 * 128 - 1).Block.HasFreeSectors);
            try
            {
                Assert.AreEqual(false, fs.GetBATBlockAndIndex(109 * 128).Block.HasFreeSectors);
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
            Assert.AreEqual(109, header.BATCount);
            Assert.AreEqual(0, header.XBATCount);

            free = fs.GetFreeBlock();
            Assert.AreEqual(false, fs.GetBATBlockAndIndex(109 * 128 - 1).Block.HasFreeSectors);
            Assert.AreEqual(true, fs.GetBATBlockAndIndex(110 * 128 - 1).Block.HasFreeSectors);
            try
            {
                Assert.AreEqual(false, fs.GetBATBlockAndIndex(110 * 128).Block.HasFreeSectors);
                Assert.Fail("Should only be 110 BATs");
            }
            //catch (IndexOutOfRangeException)
            catch (ArgumentOutOfRangeException)
            {
            }

            assertBATCount(fs, 110, 1);

            header = WriteOutAndReadHeader(fs);
            Assert.AreEqual(110, header.BATCount);
            Assert.AreEqual(1, header.XBATCount);

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
            Assert.AreEqual(false, fs.GetBATBlockAndIndex(236 * 128 - 1).Block.HasFreeSectors);
            try
            {
                Assert.AreEqual(false, fs.GetBATBlockAndIndex(236 * 128).Block.HasFreeSectors);
                Assert.Fail("Should only be 236 BATs");
            }
            catch (ArgumentOutOfRangeException)
            {
            }
            assertBATCount(fs, 236, 1);

            // Ask for another, will get our 2nd XBAT
            free = fs.GetFreeBlock();
            Assert.AreEqual(false, fs.GetBATBlockAndIndex(236 * 128 - 1).Block.HasFreeSectors);
            Assert.AreEqual(true, fs.GetBATBlockAndIndex(237 * 128 - 1).Block.HasFreeSectors);
            try
            {
                Assert.AreEqual(false, fs.GetBATBlockAndIndex(237 * 128).Block.HasFreeSectors);
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
            MemoryStream stream = new MemoryStream();
            fs.WriteFilesystem(stream);

            
            // Check that it is seen correctly
            fs = new NPOIFSFileSystem(new MemoryStream(stream.ToArray()));
            assertBATCount(fs, 237, 2);

            Assert.AreEqual(false, fs.GetBATBlockAndIndex(236 * 128 - 1).Block.HasFreeSectors);
            Assert.AreEqual(true, fs.GetBATBlockAndIndex(237 * 128 - 1).Block.HasFreeSectors);
            try
            {
                Assert.AreEqual(false, fs.GetBATBlockAndIndex(237 * 128).Block.HasFreeSectors);
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
                Assert.AreEqual(5, root.EntryCount);

                // Check by the names
                Entry thumbnail = root.GetEntry("Thumbnail");
                Entry dsi = root.GetEntry("\u0005DocumentSummaryInformation");
                Entry si = root.GetEntry("\u0005SummaryInformation");
                Entry image = root.GetEntry("Image");
                Entry tags = root.GetEntry("Tags");

                Assert.AreEqual(false, thumbnail.IsDirectoryEntry);
                Assert.AreEqual(false, dsi.IsDirectoryEntry);
                Assert.AreEqual(false, si.IsDirectoryEntry);
                Assert.AreEqual(true, image.IsDirectoryEntry);
                Assert.AreEqual(false, tags.IsDirectoryEntry);

                // Check via the iterator
                IEnumerator<Entry> it = root.Entries;
                it.MoveNext();
                Assert.AreEqual(thumbnail.Name, it.Current.Name);
                it.MoveNext();
                Assert.AreEqual(dsi.Name, it.Current.Name);
                it.MoveNext();
                Assert.AreEqual(si.Name, it.Current.Name);
                it.MoveNext();
                Assert.AreEqual(image.Name, it.Current.Name);
                it.MoveNext();
                Assert.AreEqual(tags.Name, it.Current.Name);

                // Look inside another
                DirectoryEntry imageD = (DirectoryEntry)image;
                Assert.AreEqual(7, imageD.EntryCount);
            }
        }

        [Test]
        public void TestGetDocumentEntry()
        {
            foreach (NPOIFSFileSystem fs in get512and4kFileAndInput())
            {
                DirectoryEntry root = fs.Root;
                Entry si = root.GetEntry("\x0005SummaryInformation");

                Assert.AreEqual(true, si.IsDocumentEntry);
                DocumentNode doc = (DocumentNode)si;

                // Check we can read it
                NDocumentInputStream inp = new NDocumentInputStream(doc);
                byte[] contents = new byte[doc.Size];
                Assert.AreEqual(doc.Size, inp.Read(contents));
                inp.Close();

                // Now try to build the property set
                //  ByteBuffer temp = inp.GetCurrentBuffer();
                inp = new NDocumentInputStream(doc);

                PropertySet ps = PropertySetFactory.Create(inp);
                SummaryInformation inf = (SummaryInformation)ps;

                // Check some bits in it
                Assert.AreEqual(null, inf.ApplicationName);
                Assert.AreEqual(null, inf.Author);
                Assert.AreEqual(null, inf.Subject);

                // Finish
                inp.Close();
            }
        }

        /**
            * Read a file, write it and read it again.
            * Then, alter+add some streams, write and read
        */
        [Test]
        public void TestReadWriteRead()
        {
            foreach (NPOIFSFileSystem fs in get512and4kFileAndInput())
            {
                // Check we can find the entries we expect
                // TODO Add check

                // Write out, re-load
                // TODO Add check

                // Check they're still there
                // TODO Add check

                // Check the first few and last few bytes of a few
                // TODO Add check

                // Add a test mini stream
                // TODO Add check

                // Write out, re-load
                // TODO Add check

                // Check old and new are there
                // TODO Add check

                // Add a full stream, delete a full stream
                // TODO Add check

                // Write out, re-load
                // TODO Add check

                // Check it's all there
                // TODO Add check

                // All done
                fs.Close();
            }
        }

        /**
             * Create a new file, write it and read it again
             * Then, add some streams, write and read
        */
        [Test]
        public void TestCreateWriteRead()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem();

            // Initially has a BAT but not SBAT
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(1));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(2));

            // Check that the SBAT is empty
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.Root.Property.StartBlock);

            // Write and read it
            MemoryStream baos = new MemoryStream();
            fs.WriteFilesystem(baos);
            byte[] temp = baos.ToArray();
            fs = new NPOIFSFileSystem(new MemoryStream(temp));
            //fs = new NPOIFSFileSystem(new MemoryStream(baos.ToArray()));

            // Property table entries have been added to the blocks 
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(1));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(2));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(3));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.Root.Property.StartBlock);

            // Put everything within a new directory
            DirectoryEntry testDir = fs.CreateDirectory("Test Directory");

            // Add a new Normal Stream (Normal Streams minimum 4096 bytes)
            byte[] main4096 = new byte[4096];
            main4096[0] = unchecked((byte)-10);
            main4096[4095] = unchecked((byte)-11);
            testDir.CreateDocument("Normal4096", new MemoryStream(main4096));

            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(1));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(2));
            if (fs.GetBigBlockSize() == POIFSConstants.SMALLER_BIG_BLOCK_SIZE)
            {
                Assert.AreEqual(4, fs.GetNextBlock(3));
                Assert.AreEqual(5, fs.GetNextBlock(4));
                Assert.AreEqual(6, fs.GetNextBlock(5));
                Assert.AreEqual(7, fs.GetNextBlock(6));
                Assert.AreEqual(8, fs.GetNextBlock(7));
                Assert.AreEqual(9, fs.GetNextBlock(8));
                Assert.AreEqual(10, fs.GetNextBlock(9));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(10));
                Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(11));
            }
            else
            {
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(3));
                Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(4));
            }
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.Root.Property.StartBlock);


            // Add a bigger Normal Stream
            byte[] main5124 = new byte[5124];
            main5124[0] = unchecked((byte)-22);
            main5124[5123] = unchecked((byte)-33);
            testDir.CreateDocument("Normal5124", new MemoryStream(main5124));

            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(1));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(2));
            if (fs.GetBigBlockSize() == POIFSConstants.SMALLER_BIG_BLOCK_SIZE)
            {
                Assert.AreEqual(4, fs.GetNextBlock(3));
                Assert.AreEqual(5, fs.GetNextBlock(4));
                Assert.AreEqual(6, fs.GetNextBlock(5));
                Assert.AreEqual(7, fs.GetNextBlock(6));
                Assert.AreEqual(8, fs.GetNextBlock(7));
                Assert.AreEqual(9, fs.GetNextBlock(8));
                Assert.AreEqual(10, fs.GetNextBlock(9));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(10));

                Assert.AreEqual(12, fs.GetNextBlock(11));
                Assert.AreEqual(13, fs.GetNextBlock(12));
                Assert.AreEqual(14, fs.GetNextBlock(13));
                Assert.AreEqual(15, fs.GetNextBlock(14));
                Assert.AreEqual(16, fs.GetNextBlock(15));
                Assert.AreEqual(17, fs.GetNextBlock(16));
                Assert.AreEqual(18, fs.GetNextBlock(17));
                Assert.AreEqual(19, fs.GetNextBlock(18));
                Assert.AreEqual(20, fs.GetNextBlock(19));
                Assert.AreEqual(21, fs.GetNextBlock(20));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(21));
                Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(22));
            }
            else
            {
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(3));
                Assert.AreEqual(5, fs.GetNextBlock(4));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(5));
                Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(6));
            }
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.Root.Property.StartBlock);


            // Now Add a mini stream
            byte[] mini = new byte[] { 0, 1, 2, 3, 4 };
            testDir.CreateDocument("Mini", new MemoryStream(mini));

            // Mini stream will Get one block for fat + one block for data
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(1));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(2));
            if (fs.GetBigBlockSize() == POIFSConstants.SMALLER_BIG_BLOCK_SIZE)
            {
                Assert.AreEqual(4, fs.GetNextBlock(3));
                Assert.AreEqual(5, fs.GetNextBlock(4));
                Assert.AreEqual(6, fs.GetNextBlock(5));
                Assert.AreEqual(7, fs.GetNextBlock(6));
                Assert.AreEqual(8, fs.GetNextBlock(7));
                Assert.AreEqual(9, fs.GetNextBlock(8));
                Assert.AreEqual(10, fs.GetNextBlock(9));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(10));

                Assert.AreEqual(12, fs.GetNextBlock(11));
                Assert.AreEqual(13, fs.GetNextBlock(12));
                Assert.AreEqual(14, fs.GetNextBlock(13));
                Assert.AreEqual(15, fs.GetNextBlock(14));
                Assert.AreEqual(16, fs.GetNextBlock(15));
                Assert.AreEqual(17, fs.GetNextBlock(16));
                Assert.AreEqual(18, fs.GetNextBlock(17));
                Assert.AreEqual(19, fs.GetNextBlock(18));
                Assert.AreEqual(20, fs.GetNextBlock(19));
                Assert.AreEqual(21, fs.GetNextBlock(20));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(21));
                Assert.AreEqual(23, fs.GetNextBlock(22));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(23));
                Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(24));
            }
            else
            {
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(3));
                Assert.AreEqual(5, fs.GetNextBlock(4));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(5));
                Assert.AreEqual(7, fs.GetNextBlock(6));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(7));
                Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(8));
            }
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.Root.Property.StartBlock);


            // Write and read back
            // TODO
        }

        [Test]
        public void AddBeforeWrite()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem();

            // Initially has BAT + Properties but nothing else
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(1));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(2));

            // Add to the mini stream
            // TODO

            // Add to the main stream
            // TODO

            // Write, Read, check
            // TODO
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
            nfs.WriteFilesystem(bos);
            nfs.Close();

            POIFSFileSystem pfs = new POIFSFileSystem(new MemoryStream(bos.ToArray()));
            byte[] wbDataAct = IOUtils.ToByteArray(pfs.CreateDocumentInputStream("Workbook"));

            Arrays.Equals(wbDataExp, wbDataAct);
        }

        // TODO Directory/Document create/write/read/delete/change tests
    }
}
