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
using System.Collections;

using NUnit.Framework;

using NPOI.POIFS.Common;
using TestCases;
using NPOI.POIFS.FileSystem;
using NPOI.POIFS.NIO;
using System.Collections.Generic;
using NPOI.Util;

namespace TestCases.POIFS.FileSystem
{
    /// <summary>
    /// Summary description for TestNPOIFSMiniStore
    /// </summary>
    [TestFixture]
    public class TestNPOIFSMiniStore
    {
        private static POIDataSamples _inst = POIDataSamples.GetPOIFSInstance();

        public TestNPOIFSMiniStore()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        [Test]
        public void TestNextBlock()
        {
            // It's the same on 512 byte and 4096 byte block files!
            NPOIFSFileSystem fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));
            NPOIFSFileSystem fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));
            NPOIFSFileSystem fsC = new NPOIFSFileSystem(_inst.GetFile("BlockSize4096.zvi"));
            NPOIFSFileSystem fsD = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize4096.zvi"));
            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB, fsC, fsD })
            {
                NPOIFSMiniStore ministore = fs.GetMiniStore();

                // 0 -> 51 is one stream
                for (int i = 0; i < 50; i++)
                {
                    Assert.AreEqual(i + 1, ministore.GetNextBlock(i));
                }
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(50));

                // 51 -> 103 is the next
                for (int i = 51; i < 103; i++)
                {
                    Assert.AreEqual(i + 1, ministore.GetNextBlock(i));
                }
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(103));

                // Then there are 3 one block ones
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(104));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(105));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(106));

                // 107 -> 154 is the next
                for (int i = 107; i < 154; i++)
                {
                    Assert.AreEqual(i + 1, ministore.GetNextBlock(i));
                }
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(154));

                // 155 -> 160 is the next
                for (int i = 155; i < 160; i++)
                {
                    Assert.AreEqual(i + 1, ministore.GetNextBlock(i));
                }
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(160));

                // 161 -> 166 is the next
                for (int i = 161; i < 166; i++)
                {
                    Assert.AreEqual(i + 1, ministore.GetNextBlock(i));
                }
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(166));

                // 167 -> 172 is the next
                for (int i = 167; i < 172; i++)
                {
                    Assert.AreEqual(i + 1, ministore.GetNextBlock(i));
                }
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(172));

                // Now some short ones
                Assert.AreEqual(174, ministore.GetNextBlock(173));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(174));

                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(175));

                Assert.AreEqual(177, ministore.GetNextBlock(176));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(177));

                Assert.AreEqual(179, ministore.GetNextBlock(178));
                Assert.AreEqual(180, ministore.GetNextBlock(179));
                Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(180));

                // 181 onwards is free
                for (int i = 181; i < fs.GetBigBlockSizeDetails().GetBATEntriesPerBlock(); i++)
                {
                    Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(i));
                }
                fs.Close();
            }
        }


        [Test]
        public void TestGetBlock()
        {
            // It's the same on 512 byte and 4096 byte block files!
            NPOIFSFileSystem fsA = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));
            NPOIFSFileSystem fsB = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));
            NPOIFSFileSystem fsC = new NPOIFSFileSystem(_inst.GetFile("BlockSize4096.zvi"));
            NPOIFSFileSystem fsD = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize4096.zvi"));
            foreach (NPOIFSFileSystem fs in new NPOIFSFileSystem[] { fsA, fsB, fsC, fsD })
            {
                // Mini stream should be at big block zero
                Assert.AreEqual(0, fs.PropertyTable.Root.StartBlock);

                // Grab the ministore
                NPOIFSMiniStore ministore = fs.GetMiniStore();
                ByteBuffer b;

                // Runs from the start of the data section in 64 byte chungs
                b = ministore.GetBlockAt(0);
                Assert.AreEqual((byte)0x9e, b[0]);
                Assert.AreEqual((byte)0x75, b[1]);
                Assert.AreEqual((byte)0x97, b[2]);
                Assert.AreEqual((byte)0xf6, b[3]);
                Assert.AreEqual((byte)0xff, b[4]);
                Assert.AreEqual((byte)0x21, b[5]);
                Assert.AreEqual((byte)0xd2, b[6]);
                Assert.AreEqual((byte)0x11, b[7]);

                // And the next block
                b = ministore.GetBlockAt(1);
                Assert.AreEqual((byte)0x00, b[0]);
                Assert.AreEqual((byte)0x00, b[1]);
                Assert.AreEqual((byte)0x03, b[2]);
                Assert.AreEqual((byte)0x00, b[3]);
                Assert.AreEqual((byte)0x12, b[4]);
                Assert.AreEqual((byte)0x02, b[5]);
                Assert.AreEqual((byte)0x00, b[6]);
                Assert.AreEqual((byte)0x00, b[7]);

                // Check the last data block
                b = ministore.GetBlockAt(180);
                Assert.AreEqual((byte)0x30, b[0]);
                Assert.AreEqual((byte)0x00, b[1]);
                Assert.AreEqual((byte)0x00, b[2]);
                Assert.AreEqual((byte)0x00, b[3]);
                Assert.AreEqual((byte)0x00, b[4]);
                Assert.AreEqual((byte)0x00, b[5]);
                Assert.AreEqual((byte)0x00, b[6]);
                Assert.AreEqual((byte)0x80, b[7]);

                // And the rest until the end of the big block is zeros
                for (int i = 181; i < 184; i++)
                {
                    b = ministore.GetBlockAt(i);
                    Assert.AreEqual((byte)0, b[0]);
                    Assert.AreEqual((byte)0, b[1]);
                    Assert.AreEqual((byte)0, b[2]);
                    Assert.AreEqual((byte)0, b[3]);
                    Assert.AreEqual((byte)0, b[4]);
                    Assert.AreEqual((byte)0, b[5]);
                    Assert.AreEqual((byte)0, b[6]);
                    Assert.AreEqual((byte)0, b[7]);
                }

                fs.Close();
            }
        }

        [Test]
        public void TestGetFreeBlockWithSpare()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));
            NPOIFSMiniStore ministore = fs.GetMiniStore();

            // Our 2nd SBAT block has spares
            Assert.AreEqual(false, ministore.GetBATBlockAndIndex(0).Block.HasFreeSectors);
            Assert.AreEqual(true, ministore.GetBATBlockAndIndex(128).Block.HasFreeSectors);

            // First free one at 181
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(181));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(182));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(183));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(184));

            // Ask, will get 181
            Assert.AreEqual(181, ministore.GetFreeBlock());

            // Ask again, will still get 181 as not written to
            Assert.AreEqual(181, ministore.GetFreeBlock());

            // Allocate it, then ask again
            ministore.SetNextBlock(181, POIFSConstants.END_OF_CHAIN);
            Assert.AreEqual(182, ministore.GetFreeBlock());

            fs.Close();
        }

        [Test]
        public void TestGetFreeBlockWithNonSpare()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));
            NPOIFSMiniStore ministore = fs.GetMiniStore();

            // We've spare ones from 181 to 255
            for (int i = 181; i < 256; i++)
            {
                Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(i));
            }

            // Check our SBAT free stuff is correct
            Assert.AreEqual(false, ministore.GetBATBlockAndIndex(0).Block.HasFreeSectors);
            Assert.AreEqual(true, ministore.GetBATBlockAndIndex(128).Block.HasFreeSectors);

            // Allocate all the spare ones
            for (int i = 181; i < 256; i++)
            {
                ministore.SetNextBlock(i, POIFSConstants.END_OF_CHAIN);
            }

            // SBAT are now full, but there's only the two
            Assert.AreEqual(false, ministore.GetBATBlockAndIndex(0).Block.HasFreeSectors);
            Assert.AreEqual(false, ministore.GetBATBlockAndIndex(128).Block.HasFreeSectors);
            try
            {
                Assert.AreEqual(false, ministore.GetBATBlockAndIndex(256).Block.HasFreeSectors);
                Assert.Fail("Should only be two SBATs");
            }
            catch (ArgumentOutOfRangeException) { }

            // Now ask for a free one, will need to extend the SBAT chain
            Assert.AreEqual(256, ministore.GetFreeBlock());

            Assert.AreEqual(false, ministore.GetBATBlockAndIndex(0).Block.HasFreeSectors);
            Assert.AreEqual(false, ministore.GetBATBlockAndIndex(128).Block.HasFreeSectors);
            Assert.AreEqual(true, ministore.GetBATBlockAndIndex(256).Block.HasFreeSectors);
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(254)); // 2nd SBAT 
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(255)); // 2nd SBAT
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(256)); // 3rd SBAT
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(257)); // 3rd SBAT

            fs.Close();
        }

        [Test]
        public void TestCreateBlockIfNeeded()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));
            NPOIFSMiniStore ministore = fs.GetMiniStore();

            // 178 -> 179 -> 180, 181+ is free
            Assert.AreEqual(179, ministore.GetNextBlock(178));
            Assert.AreEqual(180, ministore.GetNextBlock(179));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(180));
            for (int i = 181; i < 256; i++)
            {
                Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(i));
            }

            // However, the ministore data only covers blocks to 183
            for (int i = 0; i <= 183; i++)
            {
                ministore.GetBlockAt(i);
            }
            try
            {
                ministore.GetBlockAt(184);
                Assert.Fail("No block at 184");
            }
            catch (IndexOutOfRangeException) { }

            // The ministore itself is made up of 23 big blocks
            IEnumerator<ByteBuffer> it = new NPOIFSStream(fs, fs.Root.Property.StartBlock).GetBlockIterator();
            int count = 0;

            while (it.MoveNext())
            {
                count++;
                //it.MoveNext();
            }
            Assert.AreEqual(23, count);

            // Ask it to get block 184 with creating, it will do
            ministore.CreateBlockIfNeeded(184);

            // The ministore should be one big block bigger now
            it = new NPOIFSStream(fs, fs.Root.Property.StartBlock).GetBlockIterator();
            count = 0;
            while (it.MoveNext())
            {
                count++;
                //it.MoveNext();
            }
            Assert.AreEqual(24, count);

            // The mini block block counts now run to 191
            for (int i = 0; i <= 191; i++)
            {
                ministore.GetBlockAt(i);
            }
            try
            {
                ministore.GetBlockAt(192);
                Assert.Fail("No block at 192");
            }
            catch (IndexOutOfRangeException) { }


            // Now try writing through to 192, check that the SBAT and blocks are there
            byte[] data = new byte[15 * 64];
            NPOIFSStream stream = new NPOIFSStream(ministore, 178);
            stream.UpdateContents(data);

            // Check now
            Assert.AreEqual(179, ministore.GetNextBlock(178));
            Assert.AreEqual(180, ministore.GetNextBlock(179));
            Assert.AreEqual(181, ministore.GetNextBlock(180));
            Assert.AreEqual(182, ministore.GetNextBlock(181));
            Assert.AreEqual(183, ministore.GetNextBlock(182));
            Assert.AreEqual(184, ministore.GetNextBlock(183));
            Assert.AreEqual(185, ministore.GetNextBlock(184));
            Assert.AreEqual(186, ministore.GetNextBlock(185));
            Assert.AreEqual(187, ministore.GetNextBlock(186));
            Assert.AreEqual(188, ministore.GetNextBlock(187));
            Assert.AreEqual(189, ministore.GetNextBlock(188));
            Assert.AreEqual(190, ministore.GetNextBlock(189));
            Assert.AreEqual(191, ministore.GetNextBlock(190));
            Assert.AreEqual(192, ministore.GetNextBlock(191));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(192));
            for (int i = 193; i < 256; i++)
            {
                Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(i));
            }

            fs.Close();
        }

    }
}
