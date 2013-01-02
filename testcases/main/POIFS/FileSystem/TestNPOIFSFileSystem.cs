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
                catch(ArgumentOutOfRangeException)
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
                catch(ArgumentOutOfRangeException)
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
                catch(ArgumentOutOfRangeException)
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
                catch(ArgumentOutOfRangeException)
                {
                }

                for(int i = 0; i < 50; i++)
                    Assert.AreEqual(i+1, miniStore.GetNextBlock(i));

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

                for(int i = 22; i < 88; i++)
                Assert.AreEqual(i+1, fs.GetNextBlock(i));
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
            catch(ArgumentOutOfRangeException)
            {
            }

            // Now ask for a free one, will need to extend the file

            Assert.AreEqual(129, fs.GetFreeBlock());

            Assert.AreEqual(false, fs.GetBATBlockAndIndex(0).Block.HasFreeSectors);
            Assert.AreEqual(true, fs.GetBATBlockAndIndex(128).Block.HasFreeSectors);
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(128));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(129));

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
            catch(ArgumentOutOfRangeException)
            {
            }

            free = fs.GetFreeBlock();
            Assert.AreEqual(false, fs.GetBATBlockAndIndex(109 * 128 - 1).Block.HasFreeSectors);
            Assert.AreEqual(true, fs.GetBATBlockAndIndex(110 * 128 - 1).Block.HasFreeSectors);
            try
            {
                Assert.AreEqual(false, fs.GetBATBlockAndIndex(110 * 128).Block.HasFreeSectors);
                Assert.Fail("Should only be 110 BATs");
            }
            //catch (IndexOutOfRangeException)
            catch(ArgumentOutOfRangeException)
            {
            }

            for (int i = 109; i < 109 + 127; i++)
            {
                fs.GetFreeBlock();
                int startOffset = i * 128;
                while (fs.GetBATBlockAndIndex(startOffset).Block.HasFreeSectors)
                {
                    free = fs.GetFreeBlock();
                    fs.SetNextBlock(free, POIFSConstants.END_OF_CHAIN);
                }
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

            // Check the counts
            int numBATs = 0;
            int numXBATs = 0;
            for (int i = 0; i < 237 * 128; i++)
            {
                if (fs.GetNextBlock(i) == POIFSConstants.FAT_SECTOR_BLOCK)
                {
                    numBATs++;
                }
                if (fs.GetNextBlock(i) == POIFSConstants.DIFAT_SECTOR_BLOCK)
                {
                    numXBATs++;
                }
            }
            Assert.AreEqual(237, numBATs);
#if !HIDE_UNREACHABLE_CODE
            if (1 == 2)
            {
                // TODO Fix this - actual is 128
                Assert.AreEqual(2, numXBATs);
            }
#endif
            MemoryStream stream = new MemoryStream();
            fs.WriteFilesystem(stream);

            HeaderBlock header = new HeaderBlock(new MemoryStream(stream.ToArray()));
            Assert.AreEqual(237, header.BATCount);
#if !HIDE_UNREACHABLE_CODE
            if (1 == 2) // TODO Fix this - actual is 128
            {                
                Assert.AreEqual(2, header.XBATCount);

                fs = new NPOIFSFileSystem(new MemoryStream(stream.ToArray()));
            }
#endif

        }

        /**
        * Test that we can correctly get the list of directory
        *  entries, and the details on the files in them
        */
        [Test]
        public void TestListEntries() 
        {

              NPOIFSFileSystem fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));
              NPOIFSFileSystem fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));
              NPOIFSFileSystem fsC = new NPOIFSFileSystem(_inst.GetFile("BlockSize4096.zvi"));
              NPOIFSFileSystem fsD = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize4096.zvi"));
              foreach(NPOIFSFileSystem fs in new NPOIFSFileSystem[] {fsA,fsB,fsC,fsD}) 
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

            NPOIFSFileSystem fsB = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));
            NPOIFSFileSystem fsA = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));
            NPOIFSFileSystem fsC = new NPOIFSFileSystem(_inst.GetFile("BlockSize4096.zvi"));
            NPOIFSFileSystem fsD = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize4096.zvi"));

            
            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB, fsC, fsD })
            {
                DirectoryEntry root = fs.Root;
                Entry si = root.GetEntry("\x0005SummaryInformation");

                Assert.AreEqual(true, si.IsDocumentEntry);
                DocumentNode doc = (DocumentNode)si;

                // Check we can read it
                NDocumentInputStream inp = new NDocumentInputStream(doc);
                byte[] contents = new byte[doc.Size];
                Assert.AreEqual(doc.Size, inp.Read(contents));
                
                // Now try to build the property set
              //  ByteBuffer temp = inp.GetCurrentBuffer();
                inp = new NDocumentInputStream(doc);

                PropertySet ps = PropertySetFactory.Create(inp);
                SummaryInformation inf = (SummaryInformation)ps;

                // Check some bits in it
                Assert.AreEqual(null, inf.ApplicationName);
                Assert.AreEqual(null, inf.Author);
                Assert.AreEqual(null, inf.Subject);
            }
        }

        /**
            * Read a file, write it and read it again.
            * Then, alter+add some streams, write and read
        */
        [Test]
        public void TestReadWriteRead()
        {
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

            // Check it's still like that
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(1));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(2));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.Root.Property.StartBlock);

            // Now add a normal stream and a mini stream
            // TODO

            // TODO The rest of the test
        }

    }


}
