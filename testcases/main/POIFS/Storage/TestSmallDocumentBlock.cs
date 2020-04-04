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
using System.Collections.Generic;

namespace TestCases.POIFS.Storage
{
    /**
     * Class to Test SmallDocumentBlock functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestSmallDocumentBlock
    {
        private static byte[] testData;
        private static int testDataSize = 2999;


        public TestSmallDocumentBlock()
        {
            testData = new byte[testDataSize];
            for (int i = 0; i < testData.Length; i++)
            {
                testData[i] = (byte)i;
            }
        }

        /**
         * Test conversion from DocumentBlocks
         *
         * @exception IOException
         */
        [Test]
        public void TestConvert1()
        {
            MemoryStream stream = new MemoryStream(testData);
            List<DocumentBlock> documents = new List<DocumentBlock>();

            while (true)
            {
                DocumentBlock block = new DocumentBlock(stream, POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);

                documents.Add(block);
                if (block.PartiallyRead)
                {
                    break;
                }
            }
            SmallDocumentBlock[] results =
                SmallDocumentBlock.Convert(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, 
                    documents.ToArray(), testDataSize);

            Assert.AreEqual((testDataSize + 63) / 64, results.Length, "checking correct result size: ");
            MemoryStream output = new MemoryStream();

            for (int j = 0; j < results.Length; j++)
            {
                results[j].WriteBlocks(output);
            }
            byte[] output_array = output.ToArray();

            Assert.AreEqual(64 * results.Length,
                         output_array.Length, "checking correct output size: ");
            int index = 0;

            for (; index < testDataSize; index++)
            {
                Assert.AreEqual(testData[index],
                             output_array[index], "checking output " + index);
            }
            for (; index < output_array.Length; index++)
            {
                Assert.AreEqual((byte)0xff,
                             output_array[index], "checking output " + index);
            }
        }

        /**
         * Test conversion from byte array
         *
         * @exception IOException;
         *
         * @exception IOException
         */
        [Test]
        public void TestConvert2()
        {
            for (int j = 0; j < 320; j++)
            {
                byte[] array = new byte[j];

                for (int k = 0; k < j; k++)
                {
                    array[k] = (byte)255;       //Tony Qu changed
                }
                SmallDocumentBlock[] blocks = SmallDocumentBlock.Convert(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, array, 319);

                Assert.AreEqual(5, blocks.Length);
                MemoryStream stream = new MemoryStream();

                for (int k = 0; k < blocks.Length; k++)
                {
                    blocks[k].WriteBlocks(stream);
                }
                stream.Close();
                byte[] output = stream.ToArray();

                for (int k = 0; k < array.Length; k++)
                {
                    Assert.AreEqual(array[k], output[k], k.ToString());
                }
                for (int k = array.Length; k < 320; k++)
                {
                    Assert.AreEqual((byte)0xFF, output[k], k.ToString());
                }
            }
        }

        /**
         * Test Read method
         *
         * @exception IOException
         */
        [Test]
        public void TestRead()
        {
            MemoryStream stream = new MemoryStream(testData);
            ArrayList documents = new ArrayList();

            while (true)
            {
                DocumentBlock block = new DocumentBlock(stream, POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);

                documents.Add(block);
                if (block.PartiallyRead)
                {
                    break;
                }
            }
            SmallDocumentBlock[] blocks =
                SmallDocumentBlock.Convert(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, (BlockWritable[])documents.ToArray(typeof(DocumentBlock)), testDataSize);

            for (int j = 1; j <= testDataSize; j += 38)
            {
                byte[] buffer = new byte[j];
                int offset = 0;

                for (int k = 0; k < (testDataSize / j); k++)
                {
                    SmallDocumentBlock.Read(blocks, buffer, offset);
                    for (int n = 0; n < buffer.Length; n++)
                    {
                        Assert.AreEqual(testData[(k * j) + n], buffer[n],
                            "checking byte " + (k * j) + n);
                    }
                    offset += j;
                }
            }
        }

        /**
         * Test fill
         *
         * @exception IOException
         */
        [Test]
        public void TestFill()
        {
            for (int j = 0; j <= 8; j++)
            {
                List<SmallDocumentBlock> blocks = new List<SmallDocumentBlock>();

                for (int k = 0; k < j; k++)
                {
                    blocks.Add(null);
                }
                int result = SmallDocumentBlock.Fill(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, blocks);

                Assert.AreEqual((j + 7) / 8, result, "correct big block count: ");
                Assert.AreEqual(8 * result,
                             blocks.Count, "correct small block count: ");
                for (int m = j; m < blocks.Count; m++)
                {
                    BlockWritable block = blocks[m];
                    MemoryStream stream = new MemoryStream();

                    block.WriteBlocks(stream);
                    byte[] output = stream.ToArray();

                    Assert.AreEqual(64, output.Length, "correct output size (block[ " + m + " ]): ");
                    for (int n = 0; n < 64; n++)
                    {
                        Assert.AreEqual((byte)0xff, output[n], "correct value (block[ " + m + " ][ " + n
                                     + " ]): ");
                    }
                }
            }
        }

        /**
         * Test calcSize
         */
        [Test]
        public void TestCalcSize()
        {
            for (int j = 0; j < 10; j++)
            {
                Assert.AreEqual(j * 64,
                             SmallDocumentBlock.CalcSize(j), "testing " + j);
            }
        }

        /**
         * Test extract method
         *
         * @exception IOException
         */
        [Test]
        public void TestExtract()
        {
            byte[] data = new byte[512];
            int offset = 0;

            for (int j = 0; j < 8; j++)
            {
                for (int k = 0; k < 64; k++)
                {
                    data[offset++] = (byte)(k + j);
                }
            }
            RawDataBlock[] blocks =
            {
                new RawDataBlock(new MemoryStream(data))
            };
            IList<SmallDocumentBlock> output = SmallDocumentBlock.Extract(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, (ListManagedBlock[])blocks);

            offset = 0;
            foreach (SmallDocumentBlock block in output)
            {
                byte[] out_data = block.Data;

                Assert.AreEqual(64,
                             out_data.Length, "testing block at offset " + offset);
                for (int j = 0; j < out_data.Length; j++)
                {
                    Assert.AreEqual(data[offset], out_data[j], "testing byte at offset " + offset);
                    offset++;
                }
            }
        }

    }
}