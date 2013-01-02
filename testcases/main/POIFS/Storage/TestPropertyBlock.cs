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
using System.Collections.Generic;

using NUnit.Framework;
using NPOI.POIFS.Storage;
using NPOI.Util;
using NPOI.POIFS.FileSystem;
using NPOI.POIFS.Common;
using NPOI.POIFS.Properties;

namespace TestCases.POIFS.Storage
{
    /**
     * Class to Test PropertyBlock functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestPropertyBlock
    {

        /**
         * Constructor TestPropertyBlock
         *
         * @param name
         */

        public TestPropertyBlock()
        {
        }

        /**
         * Test constructing PropertyBlocks
         *
         * @exception IOException
         */
        [Test]
        public void TestCreatePropertyBlocks()
        {

            // Test with 0 properties
            List<NPOI.POIFS.Properties.Property> properties = new List<NPOI.POIFS.Properties.Property>();
            BlockWritable[] blocks =
                PropertyBlock.CreatePropertyBlockArray(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, properties);

            Assert.AreEqual(0, blocks.Length);

            // Test with 1 property
            properties.Add(new LocalProperty("Root Entry"));
            blocks = PropertyBlock.CreatePropertyBlockArray(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, properties);
            Assert.AreEqual(1, blocks.Length);
            byte[] testblock = new byte[512];

            for (int j = 0; j < 4; j++)
            {
                SetDefaultBlock(testblock, j);
            }
            testblock[0x0000] = (byte)'R';
            testblock[0x0002] = (byte)'o';
            testblock[0x0004] = (byte)'o';
            testblock[0x0006] = (byte)'t';
            testblock[0x0008] = (byte)' ';
            testblock[0x000A] = (byte)'E';
            testblock[0x000C] = (byte)'n';
            testblock[0x000E] = (byte)'t';
            testblock[0x0010] = (byte)'r';
            testblock[0x0012] = (byte)'y';
            testblock[0x0040] = (byte)22;
            verifyCorrect(blocks, testblock);

            // Test with 3 properties
            properties.Add(new LocalProperty("workbook"));
            properties.Add(new LocalProperty("summary"));
            blocks = PropertyBlock.CreatePropertyBlockArray(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, properties);
            Assert.AreEqual(1, blocks.Length);
            testblock[0x0080] = (byte)'w';
            testblock[0x0082] = (byte)'o';
            testblock[0x0084] = (byte)'r';
            testblock[0x0086] = (byte)'k';
            testblock[0x0088] = (byte)'b';
            testblock[0x008A] = (byte)'o';
            testblock[0x008C] = (byte)'o';
            testblock[0x008E] = (byte)'k';
            testblock[0x00C0] = (byte)18;
            testblock[0x0100] = (byte)'s';
            testblock[0x0102] = (byte)'u';
            testblock[0x0104] = (byte)'m';
            testblock[0x0106] = (byte)'m';
            testblock[0x0108] = (byte)'a';
            testblock[0x010A] = (byte)'r';
            testblock[0x010C] = (byte)'y';
            testblock[0x0140] = (byte)16;
            verifyCorrect(blocks, testblock);

            // Test with 4 properties
            properties.Add(new LocalProperty("wintery"));
            blocks = PropertyBlock.CreatePropertyBlockArray(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, properties);
            Assert.AreEqual(1, blocks.Length);
            testblock[0x0180] = (byte)'w';
            testblock[0x0182] = (byte)'i';
            testblock[0x0184] = (byte)'n';
            testblock[0x0186] = (byte)'t';
            testblock[0x0188] = (byte)'e';
            testblock[0x018A] = (byte)'r';
            testblock[0x018C] = (byte)'y';
            testblock[0x01C0] = (byte)16;
            verifyCorrect(blocks, testblock);

            // Test with 5 properties
            properties.Add(new LocalProperty("foo"));
            blocks = PropertyBlock.CreatePropertyBlockArray(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, properties);
            Assert.AreEqual(2, blocks.Length);
            testblock = new byte[1024];
            for (int j = 0; j < 8; j++)
            {
                SetDefaultBlock(testblock, j);
            }
            testblock[0x0000] = (byte)'R';
            testblock[0x0002] = (byte)'o';
            testblock[0x0004] = (byte)'o';
            testblock[0x0006] = (byte)'t';
            testblock[0x0008] = (byte)' ';
            testblock[0x000A] = (byte)'E';
            testblock[0x000C] = (byte)'n';
            testblock[0x000E] = (byte)'t';
            testblock[0x0010] = (byte)'r';
            testblock[0x0012] = (byte)'y';
            testblock[0x0040] = (byte)22;
            testblock[0x0080] = (byte)'w';
            testblock[0x0082] = (byte)'o';
            testblock[0x0084] = (byte)'r';
            testblock[0x0086] = (byte)'k';
            testblock[0x0088] = (byte)'b';
            testblock[0x008A] = (byte)'o';
            testblock[0x008C] = (byte)'o';
            testblock[0x008E] = (byte)'k';
            testblock[0x00C0] = (byte)18;
            testblock[0x0100] = (byte)'s';
            testblock[0x0102] = (byte)'u';
            testblock[0x0104] = (byte)'m';
            testblock[0x0106] = (byte)'m';
            testblock[0x0108] = (byte)'a';
            testblock[0x010A] = (byte)'r';
            testblock[0x010C] = (byte)'y';
            testblock[0x0140] = (byte)16;
            testblock[0x0180] = (byte)'w';
            testblock[0x0182] = (byte)'i';
            testblock[0x0184] = (byte)'n';
            testblock[0x0186] = (byte)'t';
            testblock[0x0188] = (byte)'e';
            testblock[0x018A] = (byte)'r';
            testblock[0x018C] = (byte)'y';
            testblock[0x01C0] = (byte)16;
            testblock[0x0200] = (byte)'f';
            testblock[0x0202] = (byte)'o';
            testblock[0x0204] = (byte)'o';
            testblock[0x0240] = (byte)8;
            verifyCorrect(blocks, testblock);
        }

        private void SetDefaultBlock(byte[] testblock, int j)
        {
            int base1 = j * 128;
            int index = 0;

            for (; index < 0x40; index++)
            {
                testblock[base1++] = (byte)0;
            }
            testblock[base1++] = (byte)2;
            testblock[base1++] = (byte)0;
            index += 2;
            for (; index < 0x44; index++)
            {
                testblock[base1++] = (byte)0;
            }
            for (; index < 0x50; index++)
            {
                testblock[base1++] = (byte)0xff;
            }
            for (; index < 0x80; index++)
            {
                testblock[base1++] = (byte)0;
            }
        }

        private void verifyCorrect(BlockWritable[] blocks, byte[] testblock)
        {
            MemoryStream stream = new MemoryStream(512
                                               * blocks.Length);

            for (int j = 0; j < blocks.Length; j++)
            {
                blocks[j].WriteBlocks(stream);
            }
            byte[] output = stream.ToArray();

            Assert.AreEqual(testblock.Length, output.Length);
            for (int j = 0; j < testblock.Length; j++)
            {
                Assert.AreEqual(testblock[j],
                             output[j], "mismatch at offset " + j);
            }
        }
    }
}