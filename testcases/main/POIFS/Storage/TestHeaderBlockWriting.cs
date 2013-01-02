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

namespace TestCases.POIFS.Storage
{
    using System;
    using System.IO;
    using System.Collections;

    using NUnit.Framework;
    using NPOI.POIFS.Storage;
    using NPOI.Util;
    using NPOI.POIFS.FileSystem;
    using NPOI.POIFS.Common;
    /**
     * Class to Test HeaderBlockWriter functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestHeaderBlockWriting
    {

        private static void ConfirmEqual(string[] expectedDataHexDumpLines, byte[] actual)
        {
            byte[] expected = RawDataUtil.Decode(expectedDataHexDumpLines);

            Assert.AreEqual(expected.Length, actual.Length);
            for (int i = 0; i < expected.Length; i++)
                Assert.AreEqual(expected[i], actual[i], "Test byte " + i);
        }

        /**
         * Test creating a HeaderBlockWriter
         *
         * @exception IOException
         */
        [Test]
        public void TestConstructors()
        {
            //
            // TODO: Add test logic	here
            //

            HeaderBlockWriter block = new HeaderBlockWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            MemoryStream output = new MemoryStream(512);

            block.WriteBlocks(output);
            byte[] copy = output.ToArray();
            string[] expected = {
             "D0 CF 11 E0 A1 B1 1A E1 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 3B 00 03 00 FE FF 09 00",
			"06 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 FE FF FF FF 00 00 00 00 00 10 00 00 FE FF FF FF",
			"00 00 00 00 FE FF FF FF 00 00 00 00 FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
        };
            ConfirmEqual(expected, copy);
            block.PropertyStart = unchecked((int)0x87654321);
            output = new MemoryStream(512);
            block.WriteBlocks(output);
            Assert.AreEqual(unchecked((int)0x87654321),  new HeaderBlockReader(new MemoryStream(output.ToArray())).PropertyStart);

        }

        /**
         * Test Setting the SBAT start block
         *
         * @exception IOException
         */
        [Test]
        public void TestSetBATStart()
        {
            HeaderBlockWriter block = new HeaderBlockWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);

            block.SBAStart = 0x1234567;
            MemoryStream output = new MemoryStream(512);

            block.WriteBlocks(output);
            byte[] copy = output.ToArray();
            string[] expected = 
        {
                "D0 CF 11 E0 A1 B1 1A E1 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 3B 00 03 00 FE FF 09 00",
                "06 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 FE FF FF FF 00 00 00 00 00 10 00 00 67 45 23 01",
                "00 00 00 00 FE FF FF FF 00 00 00 00 FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
	    		"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
	    		"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
		    	"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
		    	"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			    "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
    			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
    			"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
	    		"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
	    		"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
	    		"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
	    		"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
		    	"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
		    	"FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
        };
            ConfirmEqual(expected, copy);
        }

        /**
         * Test SetPropertyStart and GetPropertyStart
         *
         * @exception IOException
         */
        [Test]
        public void TestSetPropertyStart()
        {
            HeaderBlockWriter block = new HeaderBlockWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);

            block.PropertyStart = 0x01234567;
            MemoryStream output = new MemoryStream(512);

            block.WriteBlocks(output);
            byte[] copy = output.ToArray();

            string[] expected = 
        {
			    "D0 CF 11 E0 A1 B1 1A E1 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 3B 00 03 00 FE FF 09 00",
			    "06 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 67 45 23 01 00 00 00 00 00 10 00 00 FE FF FF FF",
			    "00 00 00 00 FE FF FF FF 00 00 00 00 FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			    "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			    "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			    "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			    "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			    "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			    "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			    "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			    "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			    "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			    "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			    "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			    "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
			    "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
        };

            ConfirmEqual(expected, copy);
        }

        /**
         * Test Setting the BAT blocks; also Tests GetBATCount,
         * GetBATArray, GetXBATCount
         *
         * @exception IOException
         */
        [Test]
        public void TestSetBATBlocks()
        {
            HeaderBlockWriter block = new HeaderBlockWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            BATBlock[] xbats = block.SetBATBlocks(5, 0x01234567);

            Assert.AreEqual(0, xbats.Length);
            Assert.AreEqual(0, HeaderBlockWriter.CalculateXBATStorageRequirements(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, 5));
            MemoryStream output = new MemoryStream(512);

            block.WriteBlocks(output);
            byte[] copy = output.ToArray();
            string[] expected = 
        {
                "D0 CF 11 E0 A1 B1 1A E1 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 3B 00 03 00 FE FF 09 00",
                "06 00 00 00 00 00 00 00 00 00 00 00 05 00 00 00 FE FF FF FF 00 00 00 00 00 10 00 00 FE FF FF FF",
                "00 00 00 00 FE FF FF FF 00 00 00 00 67 45 23 01 68 45 23 01 69 45 23 01 6A 45 23 01 6B 45 23 01",
                "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
                "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
                "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
                "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
                "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
                "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
                "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
                "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
                "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
                "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
                "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
                "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
                "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF"
        };

            ConfirmEqual(expected, copy);

            

            block = new HeaderBlockWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            xbats = block.SetBATBlocks(109, 0x01234567);
            Assert.AreEqual(0, xbats.Length);
            Assert.AreEqual(0, HeaderBlockWriter.CalculateXBATStorageRequirements(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, 109));
            output = new MemoryStream(512);
            block.WriteBlocks(output);
            copy = output.ToArray();
            string[] expected2 = 
        {
			    "D0 CF 11 E0 A1 B1 1A E1 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 3B 00 03 00 FE FF 09 00",
			    "06 00 00 00 00 00 00 00 00 00 00 00 6D 00 00 00 FE FF FF FF 00 00 00 00 00 10 00 00 FE FF FF FF",
			    "00 00 00 00 FE FF FF FF 00 00 00 00 67 45 23 01 68 45 23 01 69 45 23 01 6A 45 23 01 6B 45 23 01",
			    "6C 45 23 01 6D 45 23 01 6E 45 23 01 6F 45 23 01 70 45 23 01 71 45 23 01 72 45 23 01 73 45 23 01",
			    "74 45 23 01 75 45 23 01 76 45 23 01 77 45 23 01 78 45 23 01 79 45 23 01 7A 45 23 01 7B 45 23 01",
			    "7C 45 23 01 7D 45 23 01 7E 45 23 01 7F 45 23 01 80 45 23 01 81 45 23 01 82 45 23 01 83 45 23 01",
			    "84 45 23 01 85 45 23 01 86 45 23 01 87 45 23 01 88 45 23 01 89 45 23 01 8A 45 23 01 8B 45 23 01",
			    "8C 45 23 01 8D 45 23 01 8E 45 23 01 8F 45 23 01 90 45 23 01 91 45 23 01 92 45 23 01 93 45 23 01",
			    "94 45 23 01 95 45 23 01 96 45 23 01 97 45 23 01 98 45 23 01 99 45 23 01 9A 45 23 01 9B 45 23 01",
			    "9C 45 23 01 9D 45 23 01 9E 45 23 01 9F 45 23 01 A0 45 23 01 A1 45 23 01 A2 45 23 01 A3 45 23 01",
			    "A4 45 23 01 A5 45 23 01 A6 45 23 01 A7 45 23 01 A8 45 23 01 A9 45 23 01 AA 45 23 01 AB 45 23 01",
			    "AC 45 23 01 AD 45 23 01 AE 45 23 01 AF 45 23 01 B0 45 23 01 B1 45 23 01 B2 45 23 01 B3 45 23 01",
			    "B4 45 23 01 B5 45 23 01 B6 45 23 01 B7 45 23 01 B8 45 23 01 B9 45 23 01 BA 45 23 01 BB 45 23 01",
			    "BC 45 23 01 BD 45 23 01 BE 45 23 01 BF 45 23 01 C0 45 23 01 C1 45 23 01 C2 45 23 01 C3 45 23 01",
			    "C4 45 23 01 C5 45 23 01 C6 45 23 01 C7 45 23 01 C8 45 23 01 C9 45 23 01 CA 45 23 01 CB 45 23 01",
			    "CC 45 23 01 CD 45 23 01 CE 45 23 01 CF 45 23 01 D0 45 23 01 D1 45 23 01 D2 45 23 01 D3 45 23 01"
        };
            ConfirmEqual(expected2, copy);

            block = new HeaderBlockWriter(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            xbats = block.SetBATBlocks(256, 0x01234567);
            Assert.AreEqual(2, xbats.Length);
            Assert.AreEqual(2, HeaderBlockWriter.CalculateXBATStorageRequirements(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, 256));
            output = new MemoryStream(512);
            block.WriteBlocks(output);
            copy = output.ToArray();
            string[] expected3 = 
        {
                "D0 CF 11 E0 A1 B1 1A E1 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 3B 00 03 00 FE FF 09 00",
			    "06 00 00 00 00 00 00 00 00 00 00 00 00 01 00 00 FE FF FF FF 00 00 00 00 00 10 00 00 FE FF FF FF",
			    "00 00 00 00 67 46 23 01 02 00 00 00 67 45 23 01 68 45 23 01 69 45 23 01 6A 45 23 01 6B 45 23 01",
			    "6C 45 23 01 6D 45 23 01 6E 45 23 01 6F 45 23 01 70 45 23 01 71 45 23 01 72 45 23 01 73 45 23 01",
			    "74 45 23 01 75 45 23 01 76 45 23 01 77 45 23 01 78 45 23 01 79 45 23 01 7A 45 23 01 7B 45 23 01",
			    "7C 45 23 01 7D 45 23 01 7E 45 23 01 7F 45 23 01 80 45 23 01 81 45 23 01 82 45 23 01 83 45 23 01",
			    "84 45 23 01 85 45 23 01 86 45 23 01 87 45 23 01 88 45 23 01 89 45 23 01 8A 45 23 01 8B 45 23 01",
			    "8C 45 23 01 8D 45 23 01 8E 45 23 01 8F 45 23 01 90 45 23 01 91 45 23 01 92 45 23 01 93 45 23 01",
			    "94 45 23 01 95 45 23 01 96 45 23 01 97 45 23 01 98 45 23 01 99 45 23 01 9A 45 23 01 9B 45 23 01",
			    "9C 45 23 01 9D 45 23 01 9E 45 23 01 9F 45 23 01 A0 45 23 01 A1 45 23 01 A2 45 23 01 A3 45 23 01",
			    "A4 45 23 01 A5 45 23 01 A6 45 23 01 A7 45 23 01 A8 45 23 01 A9 45 23 01 AA 45 23 01 AB 45 23 01",
			    "AC 45 23 01 AD 45 23 01 AE 45 23 01 AF 45 23 01 B0 45 23 01 B1 45 23 01 B2 45 23 01 B3 45 23 01",
			    "B4 45 23 01 B5 45 23 01 B6 45 23 01 B7 45 23 01 B8 45 23 01 B9 45 23 01 BA 45 23 01 BB 45 23 01",
			    "BC 45 23 01 BD 45 23 01 BE 45 23 01 BF 45 23 01 C0 45 23 01 C1 45 23 01 C2 45 23 01 C3 45 23 01",
			    "C4 45 23 01 C5 45 23 01 C6 45 23 01 C7 45 23 01 C8 45 23 01 C9 45 23 01 CA 45 23 01 CB 45 23 01",
			    "CC 45 23 01 CD 45 23 01 CE 45 23 01 CF 45 23 01 D0 45 23 01 D1 45 23 01 D2 45 23 01 D3 45 23 01"
        };
            ConfirmEqual(expected3, copy);

            output = new MemoryStream(1028);
            xbats[0].WriteBlocks(output);
            xbats[1].WriteBlocks(output);
            copy = output.ToArray();
            int correct = 0x012345D4;
            int offset = 0;
            int k = 0;

            for (; k < 127; k++)
            {
                Assert.AreEqual(correct,
                             LittleEndian.GetInt(copy, offset), "XBAT entry " + k);
                correct++;
                offset += LittleEndianConsts.INT_SIZE;
            }
            Assert.AreEqual(0x01234567 + 257,
                         LittleEndian.GetInt(copy, offset), "XBAT Chain");
            offset += LittleEndianConsts.INT_SIZE;
            k++;
            for (; k < 148; k++)
            {
                Assert.AreEqual(correct,
                             LittleEndian.GetInt(copy, offset), "XBAT entry " + k);
                correct++;
                offset += LittleEndianConsts.INT_SIZE;
            }
            for (; k < 255; k++)
            {
                Assert.AreEqual(-1,
                             LittleEndian.GetInt(copy, offset), "XBAT entry " + k);
                offset += LittleEndianConsts.INT_SIZE;
            }
            Assert.AreEqual(-2,
                         LittleEndian.GetInt(copy, offset), "XBAT End of chain");
        }
    }
}