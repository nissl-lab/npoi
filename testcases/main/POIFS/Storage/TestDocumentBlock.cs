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
     * Class to Test DocumentBlock functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestDocumentBlock
    {
        private byte[] _testdata;

        public TestDocumentBlock()
        {
            _testdata = new byte[2000];
            for (int j = 0; j < _testdata.Length; j++)
            {
                _testdata[j] = (byte)j;
            }
        }

        /**
         * Test the writing DocumentBlock constructor.
         *
         * @exception IOException
         */
        [Test]
        public void TestConstructor()
        {
            MemoryStream input = new MemoryStream(_testdata);
            int index = 0;
            int size = 0;

            while (true)
            {
                byte[] data = new byte[Math.Min(_testdata.Length - index, 512)];

                Array.Copy(_testdata, index, data, 0, data.Length);
                DocumentBlock block = new DocumentBlock(input, POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);

                verifyOutput(block, data);
                size += block.Size;
                if (block.PartiallyRead)
                {
                    break;
                }
                index += 512;
            }
            Assert.AreEqual(_testdata.Length, size);
        }

        /**
         * Test static Read method
         *
         * @exception IOException
         */
        [Test]
        public void TestRead()
        {
            DocumentBlock[] blocks = new DocumentBlock[4];
            MemoryStream input = new MemoryStream(_testdata);

            for (int j = 0; j < 4; j++)
            {
                blocks[j] = new DocumentBlock(input, POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            }
            for (int j = 1; j <= 2000; j += 17)
            {
                byte[] buffer = new byte[j];
                int offset = 0;

                for (int k = 0; k < (2000 / j); k++)
                {
                    DocumentBlock.Read(blocks, buffer, offset);
                    for (int n = 0; n < buffer.Length; n++)
                    {
                        Assert.AreEqual(_testdata[(k * j) + n], buffer[n]
                            , "checking byte " + (k * j) + n);
                    }
                    offset += j;
                }
            }
        }

        /**
         * Test 'Reading' constructor
         *
         * @exception IOException
         */
        [Test]
        public void TestReadingConstructor()
        {
            RawDataBlock input =
                new RawDataBlock(new MemoryStream(_testdata));

            verifyOutput(new DocumentBlock(input), input.Data);
        }

        private void verifyOutput(DocumentBlock block, byte[] input)
        {
            Assert.AreEqual(input.Length, block.Size);
            if (input.Length < 512)
            {
                Assert.IsTrue(block.PartiallyRead);
            }
            else
            {
                Assert.IsTrue(!block.PartiallyRead);
            }
            MemoryStream output = new MemoryStream(512);

            block.WriteBlocks(output);
            byte[] copy = output.ToArray();
            int j = 0;

            for (; j < input.Length; j++)
            {
                Assert.AreEqual(input[j], copy[j]);
            }
            for (; j < 512; j++)
            {
                Assert.AreEqual((byte)0xFF, copy[j]);
            }
        }
    }
}