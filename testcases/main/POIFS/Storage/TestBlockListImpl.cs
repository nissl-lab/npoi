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
using System.IO;
using System.Collections;

using NUnit.Framework;
using NPOI.POIFS.Storage;
using NPOI.Util;
using NPOI.POIFS.FileSystem;
using NPOI.POIFS.Common;

namespace TestCases.POIFS.Storage
{
    /**
     * Class to Test BlockListImpl functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestBlockListImpl
    {

        /**
         * Test zap method
         *
         * @exception IOException
         */
        [Test]
        public void TestZap()
        {
            BlockListImpl list = new BlockListImpl();

            // verify that you can zap anything
            for (int j = -2; j < 10; j++)
            {
                list.Zap(j);
            }
            RawDataBlock[] blocks = new RawDataBlock[5];

            for (int j = 0; j < 5; j++)
            {
                blocks[j] =
                    new RawDataBlock(new MemoryStream(new byte[512]));
            }
            ListManagedBlock[] tmp = (ListManagedBlock[])blocks;

            list.SetBlocks(tmp);
            for (int j = -2; j < 10; j++)
            {
                list.Zap(j);
            }

            // verify that all blocks are gone
            for (int j = 0; j < 5; j++)
            {
                try
                {
                    list.Remove(j);
                    Assert.Fail("removing item " + j + " should not have succeeded");
                }
                catch (IOException )
                {
                }
            }
        }

        /**
         * Test Remove method
         *
         * @exception IOException
         */
        [Test]
        public void TestRemove()
        {
            BlockListImpl list = new BlockListImpl();
            RawDataBlock[] blocks = new RawDataBlock[5];
            byte[] data = new byte[512 * 5];

            for (int j = 0; j < 5; j++)
            {
                for (int k = j * 512; k < (j * 512) + 512; k++)
                {
                    data[k] = (byte)j;
                }
            }
            MemoryStream stream = new MemoryStream(data);

            for (int j = 0; j < 5; j++)
            {
                blocks[j] = new RawDataBlock(stream);
            }
            ListManagedBlock[] tmp = (ListManagedBlock[])blocks;
            list.SetBlocks(tmp);

            // verify that you can't Remove illegal indices
            for (int j = -2; j < 10; j++)
            {
                if ((j < 0) || (j >= 5))
                {
                    try
                    {
                        list.Remove(j);
                        Assert.Fail("removing item " + j + " should have Assert.Failed");
                    }
                    catch (IOException )
                    {
                    }
                }
            }

            // verify we can safely and correctly Remove all blocks
            for (int j = 0; j < 5; j++)
            {
                byte[] outPut = list.Remove(j).Data;

                for (int k = 0; k < 512; k++)
                {
                    Assert.AreEqual(data[(j * 512) + k], outPut[k], "testing block " + j + ", index " + k);
                }
            }

            // verify that all blocks are gone
            for (int j = 0; j < 5; j++)
            {
                try
                {
                    list.Remove(j);
                    Assert.Fail("removing item " + j + " should not have succeeded");
                }
                catch (IOException )
                {
                }
            }
        }

        /**
         * Test SetBAT
         *
         * @exception IOException
         */
        [Test]
        public void TestSetBAT()
        {
            BlockListImpl list = new BlockListImpl();

            list.BAT = null;
            list.BAT = new BlockAllocationTableReader(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            try
            {
                list.BAT = new BlockAllocationTableReader(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
                Assert.Fail("second attempt should have Assert.Failed");
            }
            catch (IOException )
            {
            }
        }

        /**
         * Test fetchBlocks
         *
         * @exception IOException
         */
        [Test]
        public void TestFetchBlocks()
        {

            // strategy:
            // 
            // 1. Set up a single BAT block from which to construct a
            // BAT. create nonsense blocks in the raw data block ArrayList
            // corresponding to the indices in the BAT block.
            // 2. The indices will include very short documents (0 and 1
            // block in Length), longer documents, and some screwed up
            // documents (one with a loop, one that will peek into
            // another document's data, one that includes an unused
            // document, one that includes a reserved (BAT) block, one
            // that includes a reserved (XBAT) block, and one that
            // points off into space somewhere
            BlockListImpl list = new BlockListImpl();
            ArrayList raw_blocks = new ArrayList();
            byte[] data = new byte[512];
            int offset = 0;

            LittleEndian.PutInt(data, offset, -3);   // for the BAT block itself
            offset += LittleEndianConsts.INT_SIZE;

            // document 1: Is at end of file alReady; start block = -2
            // document 2: has only one block; start block = 1
            LittleEndian.PutInt(data, offset, -2);
            offset += LittleEndianConsts.INT_SIZE;

            // document 3: has a loop in it; start block = 2
            LittleEndian.PutInt(data, offset, 2);
            offset += LittleEndianConsts.INT_SIZE;

            // document 4: peeks into document 2's data; start block = 3
            LittleEndian.PutInt(data, offset, 4);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(data, offset, 1);
            offset += LittleEndianConsts.INT_SIZE;

            // document 5: includes an unused block; start block = 5
            LittleEndian.PutInt(data, offset, 6);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(data, offset, -1);
            offset += LittleEndianConsts.INT_SIZE;

            // document 6: includes a BAT block; start block = 7
            LittleEndian.PutInt(data, offset, 8);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(data, offset, 0);
            offset += LittleEndianConsts.INT_SIZE;

            // document 7: includes an XBAT block; start block = 9
            LittleEndian.PutInt(data, offset, 10);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(data, offset, -4);
            offset += LittleEndianConsts.INT_SIZE;

            // document 8: goes off into space; start block = 11;
            LittleEndian.PutInt(data, offset, 1000);
            offset += LittleEndianConsts.INT_SIZE;

            // document 9: no screw ups; start block = 12;
            int index = 13;

            for (; offset < 508; offset += LittleEndianConsts.INT_SIZE)
            {
                LittleEndian.PutInt(data, offset, index++);
            }
            LittleEndian.PutInt(data, offset, -2);
            raw_blocks.Add(new RawDataBlock(new MemoryStream(data)));
            for (int j = raw_blocks.Count; j < 128; j++)
            {
                raw_blocks.Add(
                    new RawDataBlock(new MemoryStream(new byte[0])));
            }
            ListManagedBlock[] tmp = (ListManagedBlock[])raw_blocks.ToArray(typeof(RawDataBlock));
            list.SetBlocks(tmp);
            int[] blocks =
        {
            0
        };
            BlockAllocationTableReader table =
                new BlockAllocationTableReader(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, 1, blocks, 0, -2, list);
            int[] start_blocks =
        {
            -2, 1, 2, 3, 5, 7, 9, 11, 12
        };
            int[] expected_Length =
        {
            0, 1, -1, -1, -1, -1, -1, -1, 116
        };

            for (int j = 0; j < start_blocks.Length; j++)
            {
                try
                {
                    ListManagedBlock[] dataBlocks =
                        list.FetchBlocks(start_blocks[j],-1);

                    if (expected_Length[j] == -1)
                    {
                        Assert.Fail("document " + j + " should have Assert.Failed");
                    }
                    else
                    {
                        Assert.AreEqual(expected_Length[j], dataBlocks.Length);
                    }
                }
                catch (IOException)
                {
                    if (expected_Length[j] == -1)
                    {

                        // no problem, we expected a Assert.Failure here
                    }
                    else
                    {
                        throw;
                    }
                }
            }
        }

    }
}
