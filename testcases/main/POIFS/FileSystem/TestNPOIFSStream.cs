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
using System.Collections.Generic;
using System.IO;
using NUnit.Framework;
using NPOI.POIFS.Common;
using NPOI.POIFS.FileSystem;
using NPOI.POIFS.NIO;
using NPOI.POIFS.Storage;
using NPOI.Util;



namespace TestCases.POIFS.FileSystem
{
    /// <summary>
    /// Summary description for TestNPOIFSStream
    /// </summary>
    [TestFixture]
    public class TestNPOIFSStream
    {
        private static POIDataSamples _inst = POIDataSamples.GetPOIFSInstance();
        public TestNPOIFSStream()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        [Test]
        public void TestReadTinyStream()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));
            
            // 98 is actually the last block in a two block stream...
            NPOIFSStream stream = new NPOIFSStream(fs, 98);
            IEnumerator<ByteBuffer> i = stream.GetBlockIterator();
            Assert.AreEqual(true, i.MoveNext());
            ByteBuffer b = i.Current;
            Assert.AreEqual(false, i.MoveNext());

            // Check the contents
            Assert.AreEqual((byte)0x81, b[0]);
            Assert.AreEqual((byte)0x00, b[1]);
            Assert.AreEqual((byte)0x00, b[2]);
            Assert.AreEqual((byte)0x00, b[3]);
            Assert.AreEqual((byte)0x82, b[4]);
            Assert.AreEqual((byte)0x00, b[5]);
            Assert.AreEqual((byte)0x00, b[6]);
            Assert.AreEqual((byte)0x00, b[7]);

            fs.Close();
        }

        [Test]
        public void TestReadShortStream()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));

            // 97 -> 98 -> end
            NPOIFSStream stream = new NPOIFSStream(fs, 97);
            IEnumerator<ByteBuffer> i = stream.GetBlockIterator();

            Assert.AreEqual(true, i.MoveNext());

           // i.MoveNext();
            ByteBuffer b97 = i.Current;
            Assert.AreEqual(true, i.MoveNext());

            //i.MoveNext();
            ByteBuffer b98 = i.Current;
            Assert.AreEqual(false, i.MoveNext());

            // Check the contents of the 1st block
            Assert.AreEqual((byte)0x01, b97[0]);
            Assert.AreEqual((byte)0x00, b97[1]);
            Assert.AreEqual((byte)0x00, b97[2]);
            Assert.AreEqual((byte)0x00, b97[3]);
            Assert.AreEqual((byte)0x02, b97[4]);
            Assert.AreEqual((byte)0x00, b97[5]);
            Assert.AreEqual((byte)0x00, b97[6]);
            Assert.AreEqual((byte)0x00, b97[7]);

            // Check the contents of the 2nd block
            Assert.AreEqual((byte)0x81, b98[0]);
            Assert.AreEqual((byte)0x00, b98[1]);
            Assert.AreEqual((byte)0x00, b98[2]);
            Assert.AreEqual((byte)0x00, b98[3]);
            Assert.AreEqual((byte)0x82, b98[4]);
            Assert.AreEqual((byte)0x00, b98[5]);
            Assert.AreEqual((byte)0x00, b98[6]);
            Assert.AreEqual((byte)0x00, b98[7]);

            fs.Close();
        }

        [Test]
        public void TestReadLongerStream()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));

            ByteBuffer b0 = null;
            ByteBuffer b1 = null;
            ByteBuffer b22 = null;

            // The stream at 0 has 23 blocks in it
            NPOIFSStream stream = new NPOIFSStream(fs, 0);
            IEnumerator<ByteBuffer> i = stream.GetBlockIterator();

            int count = 0;
            while (i.MoveNext())
            {
                ByteBuffer b = i.Current;
                if (count == 0)
                {
                    b0 = b;
                }
                if (count == 1)
                {
                    b1 = b;
                }
                if (count == 22)
                {
                    b22 = b;
                }

                count++;
            }
            Assert.AreEqual(23, count);

            // Check the contents
            //  1st block is at 0
            Assert.AreEqual((byte)0x9e, b0[0]);
            Assert.AreEqual((byte)0x75, b0[1]);
            Assert.AreEqual((byte)0x97, b0[2]);
            Assert.AreEqual((byte)0xf6, b0[3]);

            //  2nd block is at 1
            Assert.AreEqual((byte)0x86, b1[0]);
            Assert.AreEqual((byte)0x09, b1[1]);
            Assert.AreEqual((byte)0x22, b1[2]);
            Assert.AreEqual((byte)0xfb, b1[3]);

            //  last block is at 89
            Assert.AreEqual((byte)0xfe, b22[0]);
            Assert.AreEqual((byte)0xff, b22[1]);
            Assert.AreEqual((byte)0x00, b22[2]);
            Assert.AreEqual((byte)0x00, b22[3]);
            Assert.AreEqual((byte)0x05, b22[4]);
            Assert.AreEqual((byte)0x01, b22[5]);
            Assert.AreEqual((byte)0x02, b22[6]);
            Assert.AreEqual((byte)0x00, b22[7]);

            fs.Close();
        }

        [Test]
        public void TestReadStream4096()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.GetFile("BlockSize4096.zvi"));

            // 0 -> 1 -> 2 -> end
            NPOIFSStream stream = new NPOIFSStream(fs, 0);
            IEnumerator<ByteBuffer> i = stream.GetBlockIterator();

            Assert.AreEqual(true, i.MoveNext());

           // i.MoveNext();
            ByteBuffer b0 = i.Current;
            Assert.AreEqual(true, i.MoveNext());

           // i.MoveNext();
            ByteBuffer b1 = i.Current;
            Assert.AreEqual(true, i.MoveNext());

           // i.MoveNext();
            ByteBuffer b2 = i.Current;
            Assert.AreEqual(false, i.MoveNext());

            // Check the contents of the 1st block
            Assert.AreEqual((byte)0x9E, b0[0]);
            Assert.AreEqual((byte)0x75, b0[1]);
            Assert.AreEqual((byte)0x97, b0[2]);
            Assert.AreEqual((byte)0xF6, b0[3]);
            Assert.AreEqual((byte)0xFF, b0[4]);
            Assert.AreEqual((byte)0x21, b0[5]);
            Assert.AreEqual((byte)0xD2, b0[6]);
            Assert.AreEqual((byte)0x11, b0[7]);

            // Check the contents of the 2nd block
            Assert.AreEqual((byte)0x00, b1[0]);
            Assert.AreEqual((byte)0x00, b1[1]);
            Assert.AreEqual((byte)0x03, b1[2]);
            Assert.AreEqual((byte)0x00, b1[3]);
            Assert.AreEqual((byte)0x00, b1[4]);
            Assert.AreEqual((byte)0x00, b1[5]);
            Assert.AreEqual((byte)0x00, b1[6]);
            Assert.AreEqual((byte)0x00, b1[7]);

            // Check the contents of the 3rd block
            Assert.AreEqual((byte)0x6D, b2[0]);
            Assert.AreEqual((byte)0x00, b2[1]);
            Assert.AreEqual((byte)0x00, b2[2]);
            Assert.AreEqual((byte)0x00, b2[3]);
            Assert.AreEqual((byte)0x03, b2[4]);
            Assert.AreEqual((byte)0x00, b2[5]);
            Assert.AreEqual((byte)0x46, b2[6]);
            Assert.AreEqual((byte)0x00, b2[7]);

            fs.Close();
        }

        [Test]
        public void TestReadFailsOnLoop()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));

            // Hack the FAT so that it goes 0->1->2->0
            fs.SetNextBlock(0, 1);
            fs.SetNextBlock(1, 2);
            fs.SetNextBlock(2, 0);

            // Now try to read
            NPOIFSStream stream = new NPOIFSStream(fs, 0);
            IEnumerator<ByteBuffer> i = stream.GetBlockIterator();
            //1st read works
            Assert.AreEqual(true, i.MoveNext());

            // 1st read works
          //  i.MoveNext();
            // 2nd read works
            Assert.AreEqual(true, i.MoveNext());

            
          // i.MoveNext();
          //  Assert.AreEqual(true, i.MoveNext());

            // 3rd read works
            //i.MoveNext();
            Assert.AreEqual(true, i.MoveNext());

            // 4th read blows up as it loops back to 0
            try
            {
                i.MoveNext();
                Assert.Fail("Loop should have been detected but wasn't!");
            }
            catch (Exception)
            {
                // Good, it was detected
            }
            //Assert.AreEqual(true, i.MoveNext());

            fs.Close();
        }

        [Test]
        public void TestReadMiniStreams()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));
            NPOIFSMiniStore ministore = fs.GetMiniStore();

            // 178 -> 179 -> 180 -> end
            NPOIFSStream stream = new NPOIFSStream(ministore, 178);
            IEnumerator<ByteBuffer> i = stream.GetBlockIterator();
            Assert.AreEqual(true, i.MoveNext());
          //  Assert.AreEqual(true, i.MoveNext());
          //  Assert.AreEqual(true, i.MoveNext());

           // i.MoveNext();
            ByteBuffer b178 = i.Current;
            Assert.AreEqual(true, i.MoveNext());
           // Assert.AreEqual(true, i.MoveNext());

           // i.MoveNext();
            ByteBuffer b179 = i.Current;
            Assert.AreEqual(true, i.MoveNext());

           // i.MoveNext();
            ByteBuffer b180 = i.Current;
            Assert.AreEqual(false, i.MoveNext());
            Assert.AreEqual(false, i.MoveNext());
           // Assert.AreEqual(false, i.MoveNext());

            // Check the contents of the 1st block
            Assert.AreEqual((byte)0xfe, b178[0]);
            Assert.AreEqual((byte)0xff, b178[1]);
            Assert.AreEqual((byte)0x00, b178[2]);
            Assert.AreEqual((byte)0x00, b178[3]);
            Assert.AreEqual((byte)0x05, b178[4]);
            Assert.AreEqual((byte)0x01, b178[5]);
            Assert.AreEqual((byte)0x02, b178[6]);
            Assert.AreEqual((byte)0x00, b178[7]);

            // And the 2nd
            Assert.AreEqual((byte)0x6c, b179[0]);
            Assert.AreEqual((byte)0x00, b179[1]);
            Assert.AreEqual((byte)0x00, b179[2]);
            Assert.AreEqual((byte)0x00, b179[3]);
            Assert.AreEqual((byte)0x28, b179[4]);
            Assert.AreEqual((byte)0x00, b179[5]);
            Assert.AreEqual((byte)0x00, b179[6]);
            Assert.AreEqual((byte)0x00, b179[7]);

            // And the 3rd
            Assert.AreEqual((byte)0x30, b180[0]);
            Assert.AreEqual((byte)0x00, b180[1]);
            Assert.AreEqual((byte)0x00, b180[2]);
            Assert.AreEqual((byte)0x00, b180[3]);
            Assert.AreEqual((byte)0x00, b180[4]);
            Assert.AreEqual((byte)0x00, b180[5]);
            Assert.AreEqual((byte)0x00, b180[6]);
            Assert.AreEqual((byte)0x80, b180[7]);

            fs.Close();
        }

        [Test]
        public void TestReplaceStream()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));

            byte[] data = new byte[512];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(i % 256);
            }

            // 98 is actually the last block in a two block stream...
            NPOIFSStream stream = new NPOIFSStream(fs, 98);
            stream.UpdateContents(data);

            // Check the reading of blocks
            IEnumerator<ByteBuffer> it = stream.GetBlockIterator();

            Assert.AreEqual(true, it.MoveNext());

          //  it.MoveNext();
            ByteBuffer b = it.Current;
            Assert.AreEqual(false, it.MoveNext());

            // Now check the contents
            data = new byte[512];
            b.Read(data);
            for (int i = 0; i < data.Length; i++)
            {
                byte exp = (byte)(i % 256);
                Assert.AreEqual(exp, data[i]);
            }

            fs.Close();
        }

        [Test]
        public void TestReplaceStreamWithLess()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));

            byte[] data = new byte[512];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(i % 256);
            }

            // 97 -> 98 -> end
            Assert.AreEqual(98, fs.GetNextBlock(97));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(98));

            // Create a 2 block stream, will become a 1 block one
            NPOIFSStream stream = new NPOIFSStream(fs, 97);
            stream.UpdateContents(data);

            // 97 should now be the end, and 98 free
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(97));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(98));

            // Check the reading of blocks
            IEnumerator<ByteBuffer> it = stream.GetBlockIterator();

            Assert.AreEqual(true, it.MoveNext());
            ByteBuffer b = it.Current;

            Assert.AreEqual(false, it.MoveNext());

            // Now check the contents
            data = new byte[512];
           // b.get(data);
            //for (int i = 0; i < b.Length; i++)
            //    data[i] = b[i];
            //Array.Copy(b, 0, data, 0, b.Length);
            b.Read(data);
            for (int i = 0; i < data.Length; i++)
            {
                byte exp = (byte)(i % 256);
                Assert.AreEqual(exp, data[i]);
            }

            fs.Close();
        }

        [Test]
        public void TestReplaceStreamWithMore()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));

            byte[] data = new byte[512 * 3];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(i % 256);
            }

            // 97 -> 98 -> end
            Assert.AreEqual(98, fs.GetNextBlock(97));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(98));

            // 100 is our first free one
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(99));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(100));

            // Create a 2 block stream, will become a 3 block one
            NPOIFSStream stream = new NPOIFSStream(fs, 97);
            stream.UpdateContents(data);

            // 97 -> 98 -> 100 -> end
            Assert.AreEqual(98, fs.GetNextBlock(97));
            Assert.AreEqual(100, fs.GetNextBlock(98));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(100));

            // Check the reading of blocks
            IEnumerator<ByteBuffer> it = stream.GetBlockIterator();
            int count = 0;
            while (it.MoveNext())
            {
                ByteBuffer b = it.Current;
                data = new byte[512];
                //b.get(data);
                //Array.Copy(b, 0, data, 0, b.Length);
                b.Read(data);

                for (int i = 0; i < data.Length; i++)
                {
                    byte exp = (byte)(i % 256);
                    Assert.AreEqual(exp, data[i]);
                }
                count++;
            }
            Assert.AreEqual(3, count);

            fs.Close();
        }

        [Test]
        public void TestWriteNewStream()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));

            // 100 is our first free one
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(99));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(100));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(101));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(102));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(103));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(104));


            // Add a single block one
            byte[] data = new byte[512];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(i % 256);
            }

            NPOIFSStream stream = new NPOIFSStream(fs);
            stream.UpdateContents(data);

            // Check it was allocated properly
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(99));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(100));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(101));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(102));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(103));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(104));

            // And check the contents
            IEnumerator<ByteBuffer> it = stream.GetBlockIterator();
            int count = 0;
            while (it.MoveNext())
            {
                ByteBuffer b = it.Current;

                data = new byte[512];
                //b.get(data);
                //Array.Copy(b, 0, data, 0, b.Length);
                b.Read(data);
                for (int i = 0; i < data.Length; i++)
                {
                    byte exp = (byte)(i % 256);
                    Assert.AreEqual(exp, data[i]);
                }
                count++;
            }
            Assert.AreEqual(1, count);


            // And a multi block one
            data = new byte[512 * 3];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(i % 256);
            }

            stream = new NPOIFSStream(fs);
            stream.UpdateContents(data);

            // Check it was allocated properly
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(99));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(100));
            Assert.AreEqual(102, fs.GetNextBlock(101));
            Assert.AreEqual(103, fs.GetNextBlock(102));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(103));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(104));

            // And check the contents
            it = stream.GetBlockIterator();
            count = 0;
            while (it.MoveNext())
            {
                ByteBuffer b = it.Current;
                data = new byte[512];
                //b.get(data);
               // Array.Copy(b, 0, data, 0, b.Length);
                b.Read(data);
                for (int i = 0; i < data.Length; i++)
                {
                    byte exp = (byte)(i % 256);
                    Assert.AreEqual(exp, data[i]);
                }
                count++;
            }
            Assert.AreEqual(3, count);

            // Free it
            stream.Free();
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(99));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(100));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(101));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(102));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(103));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(104));

            fs.Close();
        }

        [Test]
        public void TestWriteNewStreamExtraFATs()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));

            // Allocate almost all the blocks
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(99));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(100));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(127));
            for (int i = 100; i < 127; i++)
            {
                fs.SetNextBlock(i, POIFSConstants.END_OF_CHAIN);
            }
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(127));
            Assert.AreEqual(true, fs.GetBATBlockAndIndex(0).Block.HasFreeSectors);


            // Write a 3 block stream
            byte[] data = new byte[512 * 3];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(i % 256);
            }
            NPOIFSStream stream = new NPOIFSStream(fs);
            stream.UpdateContents(data);

            // Check we got another BAT
            Assert.AreEqual(false, fs.GetBATBlockAndIndex(0).Block.HasFreeSectors);
            Assert.AreEqual(true, fs.GetBATBlockAndIndex(128).Block.HasFreeSectors);

            // the BAT will be in the first spot of the new block
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(126));
            Assert.AreEqual(129, fs.GetNextBlock(127));
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(128));
            Assert.AreEqual(130, fs.GetNextBlock(129));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(130));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(131));

            fs.Close();
        }

        [Test]
        public void TestWriteStream4096()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize4096.zvi"));

            // 0 -> 1 -> 2 -> end
            Assert.AreEqual(1, fs.GetNextBlock(0));
            Assert.AreEqual(2, fs.GetNextBlock(1));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(2));
            Assert.AreEqual(4, fs.GetNextBlock(3));

            // First free one is at 15
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(14));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(15));


            // Write a 5 block file 
            byte[] data = new byte[4096 * 5];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(i % 256);
            }
            NPOIFSStream stream = new NPOIFSStream(fs, 0);
            stream.UpdateContents(data);


            // Check it
            Assert.AreEqual(1, fs.GetNextBlock(0));
            Assert.AreEqual(2, fs.GetNextBlock(1));
            Assert.AreEqual(15, fs.GetNextBlock(2)); // Jumps
            Assert.AreEqual(4, fs.GetNextBlock(3));  // Next stream
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, fs.GetNextBlock(14));
            Assert.AreEqual(16, fs.GetNextBlock(15)); // Continues
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, fs.GetNextBlock(16)); // Ends
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, fs.GetNextBlock(17)); // Free

            // Check the contents too
            IEnumerator<ByteBuffer> it = stream.GetBlockIterator();
            int count = 0;
            while (it.MoveNext())
            {
                ByteBuffer b = it.Current;
                data = new byte[512];
               // b.get(data);
              //  Array.Copy(b, 0, data, 0, b.Length);
                b.Read(data);
                for (int i = 0; i < data.Length; i++)
                {
                    byte exp = (byte)(i % 256);
                    Assert.AreEqual(exp, data[i]);
                }
                count++;
            }
            Assert.AreEqual(5, count);

            fs.Close();
        }

        [Test]
        public void TestWriteMiniStreams()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.OpenResourceAsStream("BlockSize512.zvi"));
            NPOIFSMiniStore ministore = fs.GetMiniStore();
            NPOIFSStream stream = new NPOIFSStream(ministore, 178);

            // 178 -> 179 -> 180 -> end
            Assert.AreEqual(179, ministore.GetNextBlock(178));
            Assert.AreEqual(180, ministore.GetNextBlock(179));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(180));


            // Try writing 3 full blocks worth
            byte[] data = new byte[64 * 3];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)i;
            }
            stream = new NPOIFSStream(ministore, 178);
            stream.UpdateContents(data);

            // Check
            Assert.AreEqual(179, ministore.GetNextBlock(178));
            Assert.AreEqual(180, ministore.GetNextBlock(179));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(180));

            stream = new NPOIFSStream(ministore, 178);
            IEnumerator<ByteBuffer> it = stream.GetBlockIterator();
            it.MoveNext();
            ByteBuffer b178 = it.Current;
            it.MoveNext();
            ByteBuffer b179 = it.Current;
            it.MoveNext();
            ByteBuffer b180 = it.Current;

            Assert.AreEqual(false, it.MoveNext());

            Assert.AreEqual((byte)0x00, b178.Read());
            Assert.AreEqual((byte)0x01, b178.Read());
            Assert.AreEqual((byte)0x40, b179.Read());
            Assert.AreEqual((byte)0x41, b179.Read());
            Assert.AreEqual((byte)0x80, b180.Read());
            Assert.AreEqual((byte)0x81, b180.Read());


            // Try writing just into 3 blocks worth
            data = new byte[64 * 2 + 12];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(i + 4);
            }
            stream = new NPOIFSStream(ministore, 178);
            stream.UpdateContents(data);

            // Check
            Assert.AreEqual(179, ministore.GetNextBlock(178));
            Assert.AreEqual(180, ministore.GetNextBlock(179));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(180));

            stream = new NPOIFSStream(ministore, 178);
            it = stream.GetBlockIterator();
            it.MoveNext();
            b178 = it.Current;
            it.MoveNext();
            b179 = it.Current;
            it.MoveNext();
            b180 = it.Current;
            Assert.AreEqual(false, it.MoveNext());

            Assert.AreEqual((byte)0x04, b178.Read());
            Assert.AreEqual((byte)0x05, b178.Read());
            Assert.AreEqual((byte)0x44, b179.Read());
            Assert.AreEqual((byte)0x45, b179.Read());
            Assert.AreEqual((byte)0x84, b180.Read());
            Assert.AreEqual((byte)0x85, b180.Read());


            // Try writing 1, should truncate
            data = new byte[12];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(i + 9);
            }
            stream = new NPOIFSStream(ministore, 178);
            stream.UpdateContents(data);

            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(178));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(179));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(180));

            stream = new NPOIFSStream(ministore, 178);
            it = stream.GetBlockIterator();
            it.MoveNext();
            b178 = it.Current;
            Assert.AreEqual(false, it.MoveNext());

            Assert.AreEqual((byte)0x09, b178[0]);
            Assert.AreEqual((byte)0x0a, b178[1]);

            // Try writing 5, should extend
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(178));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(179));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(180));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(181));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(182));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(183));

            data = new byte[64 * 4 + 12];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(i + 3);
            }
            stream = new NPOIFSStream(ministore, 178);
            stream.UpdateContents(data);

            Assert.AreEqual(179, ministore.GetNextBlock(178));
            Assert.AreEqual(180, ministore.GetNextBlock(179));
            Assert.AreEqual(181, ministore.GetNextBlock(180));
            Assert.AreEqual(182, ministore.GetNextBlock(181));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(182));

            stream = new NPOIFSStream(ministore, 178);
            it = stream.GetBlockIterator();
            it.MoveNext();
            b178 = it.Current;
            it.MoveNext();
            b179 = it.Current;
            it.MoveNext();
            b180 = it.Current;
            it.MoveNext();
            ByteBuffer b181 = it.Current;
            it.MoveNext();
            ByteBuffer b182 = it.Current;
            Assert.AreEqual(false, it.MoveNext());

            Assert.AreEqual((byte)0x03, b178[0]);
            Assert.AreEqual((byte)0x04, b178[1]);
            Assert.AreEqual((byte)0x43, b179[0]);
            Assert.AreEqual((byte)0x44, b179[1]);
            Assert.AreEqual((byte)0x83, b180[0]);
            Assert.AreEqual((byte)0x84, b180[1]);
            Assert.AreEqual((byte)0xc3, b181[0]);
            Assert.AreEqual((byte)0xc4, b181[1]);
            Assert.AreEqual((byte)0x03, b182[0]);
            Assert.AreEqual((byte)0x04, b182[1]);


            // Write lots, so it needs another big block
            ministore.GetBlockAt(183);
            try
            {
                ministore.GetBlockAt(184);
                Assert.Fail("Block 184 should be off the end of the list");
            }
           // catch (ArgumentOutOfRangeException e)
            catch(Exception)
            {
            }

            data = new byte[64 * 6 + 12];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(i + 1);
            }
            stream = new NPOIFSStream(ministore, 178);
            stream.UpdateContents(data);

            // Should have added 2 more blocks to the chain
            Assert.AreEqual(179, ministore.GetNextBlock(178));
            Assert.AreEqual(180, ministore.GetNextBlock(179));
            Assert.AreEqual(181, ministore.GetNextBlock(180));
            Assert.AreEqual(182, ministore.GetNextBlock(181));
            Assert.AreEqual(183, ministore.GetNextBlock(182));
            Assert.AreEqual(184, ministore.GetNextBlock(183));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, ministore.GetNextBlock(184));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, ministore.GetNextBlock(185));

            // Block 184 should exist
            ministore.GetBlockAt(183);
            ministore.GetBlockAt(184);
            ministore.GetBlockAt(185);

            // Check contents
            stream = new NPOIFSStream(ministore, 178);
            it = stream.GetBlockIterator();
            it.MoveNext();
            b178 = it.Current;
            it.MoveNext();
            b179 = it.Current;
            it.MoveNext();
            b180 = it.Current;
            it.MoveNext();
            b181 = it.Current;
            it.MoveNext();
            b182 = it.Current;
            it.MoveNext();
            ByteBuffer b183 = it.Current;
            it.MoveNext();
            ByteBuffer b184 = it.Current;
            Assert.AreEqual(false, it.MoveNext());

            Assert.AreEqual((byte)0x01, b178[0]);
            Assert.AreEqual((byte)0x02, b178[1]);
            Assert.AreEqual((byte)0x41, b179[0]);
            Assert.AreEqual((byte)0x42, b179[1]);
            Assert.AreEqual((byte)0x81, b180[0]);
            Assert.AreEqual((byte)0x82, b180[1]);
            Assert.AreEqual((byte)0xc1, b181[0]);
            Assert.AreEqual((byte)0xc2, b181[1]);
            Assert.AreEqual((byte)0x01, b182[0]);
            Assert.AreEqual((byte)0x02, b182[1]);
            Assert.AreEqual((byte)0x41, b183[0]);
            Assert.AreEqual((byte)0x42, b183[1]);
            Assert.AreEqual((byte)0x81, b184[0]);
            Assert.AreEqual((byte)0x82, b184[1]);

            fs.Close();
        }

        [Test]
        public void TestWriteFailsOnLoop()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(_inst.GetFile("BlockSize512.zvi"));

            // Hack the FAT so that it goes 0->1->2->0
            fs.SetNextBlock(0, 1);
            fs.SetNextBlock(1, 2);
            fs.SetNextBlock(2, 0);

            // Try to write a large amount, should fail on the write
            byte[] data = new byte[512 * 4];
            NPOIFSStream stream = new NPOIFSStream(fs, 0);
            try
            {
                stream.UpdateContents(data);
                Assert.Fail("Loop should have been detected but wasn't!");
            }
            catch (Exception) { }

            // Now reset, and try on a small bit
            // Should fail during the freeing set
            fs.SetNextBlock(0, 1);
            fs.SetNextBlock(1, 2);
            fs.SetNextBlock(2, 0);

            data = new byte[512];
            stream = new NPOIFSStream(fs, 0);
            try
            {
                stream.UpdateContents(data);
                Assert.Fail("Loop should have been detected but wasn't!");
            }
            catch (Exception) { }

            fs.Close();
        }

        [Test]
        public void TestReadWriteNewStream()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem();
            NPOIFSStream stream = new NPOIFSStream(fs);

            // Check our filesystem has a BAT and the Properties
            Assert.AreEqual(2, fs.GetFreeBlock());
            BATBlock bat = fs.GetBATBlockAndIndex(0).Block;
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, bat.GetValueAt(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(1));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(2));

            // Check the stream as-is
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, stream.GetStartBlock());
            try
            {
                stream.GetBlockIterator();
                Assert.Fail("Shouldn't be able to get an iterator before writing");
            }
            catch (Exception) { }

            // Write in two blocks
            byte[] data = new byte[512 + 20];
            for (int i = 0; i < 512; i++)
            {
                data[i] = (byte)(i % 256);
            }
            for (int i = 512; i < data.Length; i++)
            {
                data[i] = (byte)(i % 256 + 100);
            }
            stream.UpdateContents(data);

            // Check now
            Assert.AreEqual(4, fs.GetFreeBlock());
            bat = fs.GetBATBlockAndIndex(0).Block;
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, bat.GetValueAt(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(1));
            Assert.AreEqual(3, bat.GetValueAt(2));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(3));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(4));


            IEnumerator<ByteBuffer> it = stream.GetBlockIterator();

            Assert.AreEqual(true, it.MoveNext());
            ByteBuffer b = it.Current;

            byte[] read = new byte[512];
            //b.get(read);
           // Array.Copy(b, 0, read, 0, b.Length);
            b.Read(read);
            for (int i = 0; i < read.Length; i++)
            {
                //Assert.AreEqual("Wrong value at " + i, data[i], read[i]);
                Assert.AreEqual(data[i], read[i], "Wrong value at " + i);
            }

            Assert.AreEqual(true, it.MoveNext());
            b = it.Current;

            read = new byte[512];
            //b.get(read);
            //Array.Copy(b, 0, read, 0, b.Length);
            b.Read(read);
            for (int i = 0; i < 20; i++)
            {
                Assert.AreEqual(data[i + 512], read[i]);
            }
            for (int i = 20; i < read.Length; i++)
            {
                Assert.AreEqual(0, read[i]);
            }

            Assert.AreEqual(false, it.MoveNext());

            fs.Close();
        }
        /**
        * Writes a stream, then Replaces it
        */
        [Test]
        public void TestWriteThenReplace()
        {
            NPOIFSFileSystem fs = new NPOIFSFileSystem();

            // Starts empty
            BATBlock bat = fs.GetBATBlockAndIndex(0).Block;
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, bat.GetValueAt(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(1));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(2));

            // Write something that uses a main stream
            byte[] main4106 = new byte[4106];
            main4106[0] = unchecked((byte)-10);
            main4106[4105] = unchecked((byte)-11);
            DocumentEntry normal = fs.Root.CreateDocument(
                    "Normal", new MemoryStream(main4106));

            // Should have used 9 blocks
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, bat.GetValueAt(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(1));
            Assert.AreEqual(3, bat.GetValueAt(2));
            Assert.AreEqual(4, bat.GetValueAt(3));
            Assert.AreEqual(5, bat.GetValueAt(4));
            Assert.AreEqual(6, bat.GetValueAt(5));
            Assert.AreEqual(7, bat.GetValueAt(6));
            Assert.AreEqual(8, bat.GetValueAt(7));
            Assert.AreEqual(9, bat.GetValueAt(8));
            Assert.AreEqual(10, bat.GetValueAt(9));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(10));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(11));

            normal = (DocumentEntry)fs.Root.GetEntry("Normal");
            Assert.AreEqual(4106, normal.Size);
            Assert.AreEqual(4106, ((DocumentNode)normal).Property.Size);


            // Replace with one still big enough for a main stream, but one block smaller
            byte[] main4096 = new byte[4096];
            main4096[0] = unchecked((byte)-10);
            main4096[4095] = unchecked((byte)-11);

            NDocumentOutputStream nout = new NDocumentOutputStream(normal);
            nout.Write(main4096);
            nout.Close();

            // Will have dropped to 8
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, bat.GetValueAt(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(1));
            Assert.AreEqual(3, bat.GetValueAt(2));
            Assert.AreEqual(4, bat.GetValueAt(3));
            Assert.AreEqual(5, bat.GetValueAt(4));
            Assert.AreEqual(6, bat.GetValueAt(5));
            Assert.AreEqual(7, bat.GetValueAt(6));
            Assert.AreEqual(8, bat.GetValueAt(7));
            Assert.AreEqual(9, bat.GetValueAt(8));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(9));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(10));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(11));

            normal = (DocumentEntry)fs.Root.GetEntry("Normal");
            Assert.AreEqual(4096, normal.Size);
            Assert.AreEqual(4096, ((DocumentNode)normal).Property.Size);


            // Write and check
            fs = TestNPOIFSFileSystem.WriteOutAndReadBack(fs);
            bat = fs.GetBATBlockAndIndex(0).Block;

            // Will have properties, but otherwise the same
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, bat.GetValueAt(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(1));
            Assert.AreEqual(3, bat.GetValueAt(2));
            Assert.AreEqual(4, bat.GetValueAt(3));
            Assert.AreEqual(5, bat.GetValueAt(4));
            Assert.AreEqual(6, bat.GetValueAt(5));
            Assert.AreEqual(7, bat.GetValueAt(6));
            Assert.AreEqual(8, bat.GetValueAt(7));
            Assert.AreEqual(9, bat.GetValueAt(8));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(9)); // End of Normal
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(10)); // Props
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(11));

            normal = (DocumentEntry)fs.Root.GetEntry("Normal");
            Assert.AreEqual(4096, normal.Size);
            Assert.AreEqual(4096, ((DocumentNode)normal).Property.Size);


            // Make longer, take 1 block After the properties too
            normal = (DocumentEntry)fs.Root.GetEntry("Normal");
            nout = new NDocumentOutputStream(normal);
            nout.Write(main4106);
            nout.Close();

            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, bat.GetValueAt(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(1));
            Assert.AreEqual(3, bat.GetValueAt(2));
            Assert.AreEqual(4, bat.GetValueAt(3));
            Assert.AreEqual(5, bat.GetValueAt(4));
            Assert.AreEqual(6, bat.GetValueAt(5));
            Assert.AreEqual(7, bat.GetValueAt(6));
            Assert.AreEqual(8, bat.GetValueAt(7));
            Assert.AreEqual(9, bat.GetValueAt(8));
            Assert.AreEqual(11, bat.GetValueAt(9));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(10)); // Props
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(11)); // Normal
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(12));

            normal = (DocumentEntry)fs.Root.GetEntry("Normal");
            Assert.AreEqual(4106, normal.Size);
            Assert.AreEqual(4106, ((DocumentNode)normal).Property.Size);


            // Make it small, will trigger the SBAT stream and free lots up
            byte[] mini = new byte[] { 42, 0, 1, 2, 3, 4, 42 };
            normal = (DocumentEntry)fs.Root.GetEntry("Normal");
            nout = new NDocumentOutputStream(normal);
            nout.Write(mini);
            nout.Close();

            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, bat.GetValueAt(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(1));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(2)); // SBAT
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(3)); // Mini Stream
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(4));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(5));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(6));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(7));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(8));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(9));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(10)); // Props
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(11));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(12));

            normal = (DocumentEntry)fs.Root.GetEntry("Normal");
            Assert.AreEqual(7, normal.Size);
            Assert.AreEqual(7, ((DocumentNode)normal).Property.Size);


            // Finally back to big again
            nout = new NDocumentOutputStream(normal);
            nout.Write(main4096);
            nout.Close();

            // Will keep the mini stream, now empty
            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, bat.GetValueAt(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(1));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(2)); // SBAT
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(3)); // Mini Stream
            Assert.AreEqual(5, bat.GetValueAt(4));
            Assert.AreEqual(6, bat.GetValueAt(5));
            Assert.AreEqual(7, bat.GetValueAt(6));
            Assert.AreEqual(8, bat.GetValueAt(7));
            Assert.AreEqual(9, bat.GetValueAt(8));
            Assert.AreEqual(11, bat.GetValueAt(9));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(10)); // Props
            Assert.AreEqual(12, bat.GetValueAt(11));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(12));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(13));

            normal = (DocumentEntry)fs.Root.GetEntry("Normal");
            Assert.AreEqual(4096, normal.Size);
            Assert.AreEqual(4096, ((DocumentNode)normal).Property.Size);


            // Save, re-load, re-check
            fs = TestNPOIFSFileSystem.WriteOutAndReadBack(fs);
            bat = fs.GetBATBlockAndIndex(0).Block;

            Assert.AreEqual(POIFSConstants.FAT_SECTOR_BLOCK, bat.GetValueAt(0));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(1));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(2)); // SBAT
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(3)); // Mini Stream
            Assert.AreEqual(5, bat.GetValueAt(4));
            Assert.AreEqual(6, bat.GetValueAt(5));
            Assert.AreEqual(7, bat.GetValueAt(6));
            Assert.AreEqual(8, bat.GetValueAt(7));
            Assert.AreEqual(9, bat.GetValueAt(8));
            Assert.AreEqual(11, bat.GetValueAt(9));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(10)); // Props
            Assert.AreEqual(12, bat.GetValueAt(11));
            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, bat.GetValueAt(12));
            Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, bat.GetValueAt(13));

            normal = (DocumentEntry)fs.Root.GetEntry("Normal");
            Assert.AreEqual(4096, normal.Size);
            Assert.AreEqual(4096, ((DocumentNode)normal).Property.Size);

            fs.Close();
        }

    }
}
