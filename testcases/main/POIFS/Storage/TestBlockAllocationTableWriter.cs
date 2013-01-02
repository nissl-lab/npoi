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
using System.Text;
using System.Collections.Generic;
using System.IO;

using NUnit.Framework;

using NPOI.POIFS.Storage;
using NPOI.Util;
using NPOI.POIFS.Common;


namespace TestCases.POIFS.Storage
{
    /**
     * Class to Test BlockAllocationTableWriter functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestBlockAllocationTableWriter
    {

        /**
         * Constructor TestBlockAllocationTableWriter
         *
         * @param name
         */

        public TestBlockAllocationTableWriter()
        {
        }

        /**
         * Test the AllocateSpace method.
         */
        [Test]
        public void TestAllocateSpace()
        {
            BlockAllocationTableWriter table = new BlockAllocationTableWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);

            int[] blockSizes = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };

            int expectedIndex = 0;

            for (int i = 0; i < blockSizes.Length; i++)
            {
                Assert.AreEqual(expectedIndex, table.AllocateSpace(blockSizes[i]));
                expectedIndex += blockSizes[i];
            }
        }

        /**
         * Test the createBlocks method
         *
         * @exception IOException
         */
        [Test]
        public void TestCreateBlocks()
        {
            BlockAllocationTableWriter table = new BlockAllocationTableWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);

            table.AllocateSpace(127);
            table.CreateBlocks();
            VerifyBlocksCreated(table, 1);

            table = new BlockAllocationTableWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            table.AllocateSpace(128);
            table.CreateBlocks();
            VerifyBlocksCreated(table, 2);

            table = new BlockAllocationTableWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            table.AllocateSpace(254);
            table.CreateBlocks();
            VerifyBlocksCreated(table, 2);

            table = new BlockAllocationTableWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            table.AllocateSpace(255);
            table.CreateBlocks();
            VerifyBlocksCreated(table, 3);

            table = new BlockAllocationTableWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            table.AllocateSpace(13843);
            table.CreateBlocks();
            VerifyBlocksCreated(table, 109);

            table = new BlockAllocationTableWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            table.AllocateSpace(13844);
            table.CreateBlocks();
            VerifyBlocksCreated(table, 110);

            table = new BlockAllocationTableWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            table.AllocateSpace(13969);
            table.CreateBlocks();
            VerifyBlocksCreated(table, 110);

            table = new BlockAllocationTableWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            table.AllocateSpace(13970);
            table.CreateBlocks();
            VerifyBlocksCreated(table, 111);
        }

        /**
         * Test content produced by BlockAllocationTableWriter
         *
         * @exception IOException
         */
        [Test]
        public void TestProduct()
        {
            BlockAllocationTableWriter table = new BlockAllocationTableWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);

            for (int i = 1; i <= 22; i++)
                table.AllocateSpace(i);

            table.CreateBlocks();
            MemoryStream stream = new MemoryStream();

            table.WriteBlocks(stream);
            byte[] output = stream.ToArray();

            Assert.AreEqual(1024, output.Length);
            byte[] expected = new byte[1024];

            for (int i = 0; i < expected.Length; i++)
            {
                expected[i] = (byte)0xFF;
            }
            int offset = 0;
            int blockIndex = 1;

            for (int i = 1; i <= 22; i++)
            {
                int limit = i - 1;

                for (int j = 0; j < limit; j++)
                {
                    LittleEndian.PutInt(expected, offset, blockIndex++);
                    offset += LittleEndianConsts.INT_SIZE;
                }

                LittleEndian.PutInt(expected, offset, POIFSConstants.END_OF_CHAIN);

                offset += 4;
                blockIndex++;
            }

            LittleEndian.PutInt(expected, offset, blockIndex++);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(expected, offset, POIFSConstants.END_OF_CHAIN);

            for (int i = 0; i < expected.Length; i++)
                Assert.AreEqual(expected[i], output[i], "At offset " + i);
        }

        public static void VerifyBlocksCreated(BlockAllocationTableWriter table, int count)
        {
            MemoryStream stream = new MemoryStream();

            table.WriteBlocks(stream);
            byte[] output = stream.ToArray();

            Assert.AreEqual(count * 512, output.Length);
        }

    }
}