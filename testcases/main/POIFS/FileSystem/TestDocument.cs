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
using System.Collections;
using System.IO;

using NUnit.Framework;

using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NPOI.POIFS.Storage;
using NPOI.POIFS.Properties;

using TestCases.POIFS.FileSystem;

namespace TestCases.POIFS.FileSystem
{
    /**
     * Class to Test POIFSDocument functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestDocument
    {

        /**
         * Constructor TestDocument
         *
         * @param name
         */

        public TestDocument()
        {

        }

        /**
         * Integration Test -- really about all we can do
         *
         * @exception IOException
         */
        [Test]
        public void TestOPOIFSDocument()
        {

            // Verify correct number of blocks Get Created for document
            // that is exact multituple of block size
            OPOIFSDocument document;
            byte[] array = new byte[4096];

            for (int j = 0; j < array.Length; j++)
            {
                array[j] = (byte)j;
            }
            document = new OPOIFSDocument("foo", new SlowInputStream(new MemoryStream(array)));
            checkDocument(document, array);

            // Verify correct number of blocks Get Created for document
            // that is not an exact multiple of block size
            array = new byte[4097];
            for (int j = 0; j < array.Length; j++)
            {
                array[j] = (byte)j;
            }
            document = new OPOIFSDocument("bar", new MemoryStream(array));
            checkDocument(document, array);

            // Verify correct number of blocks Get Created for document
            // that is small
            array = new byte[4095];
            for (int j = 0; j < array.Length; j++)
            {
                array[j] = (byte)j;
            }
            document = new OPOIFSDocument("_bar", new MemoryStream(array));
            checkDocument(document, array);

            // Verify correct number of blocks Get Created for document
            // that is rather small
            array = new byte[199];
            for (int j = 0; j < array.Length; j++)
            {
                array[j] = (byte)j;
            }
            document = new OPOIFSDocument("_bar2",
                                         new MemoryStream(array));
            checkDocument(document, array);

            // Verify that output is correct
            array = new byte[4097];
            for (int j = 0; j < array.Length; j++)
            {
                array[j] = (byte)j;
            }
            document = new OPOIFSDocument("foobar",
                                         new MemoryStream(array));
            checkDocument(document, array);
            document.StartBlock=0x12345678;   // what a big file!!
            DocumentProperty property = document.DocumentProperty;
            MemoryStream stream = new MemoryStream();

            property.WriteData(stream);
            byte[] output = stream.ToArray();
            byte[] array2 =
        {
            ( byte ) 'f', ( byte ) 0, ( byte ) 'o', ( byte ) 0, ( byte ) 'o',
            ( byte ) 0, ( byte ) 'b', ( byte ) 0, ( byte ) 'a', ( byte ) 0,
            ( byte ) 'r', ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 14,
            ( byte ) 0, ( byte ) 2, ( byte ) 1, unchecked(( byte ) -1), unchecked(( byte ) -1),
            unchecked(( byte ) -1), unchecked(( byte ) -1), unchecked(( byte ) -1), unchecked(( byte ) -1), unchecked(( byte ) -1),
            unchecked(( byte ) -1), unchecked(( byte ) -1), unchecked(( byte ) -1), unchecked(( byte ) -1), unchecked(( byte ) -1),
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0x78, ( byte ) 0x56, ( byte ) 0x34,
            ( byte ) 0x12, ( byte ) 1, ( byte ) 16, ( byte ) 0, ( byte ) 0,
            ( byte ) 0, ( byte ) 0, ( byte ) 0, ( byte ) 0
        };

            Assert.AreEqual(array2.Length, output.Length);
            for (int j = 0; j < output.Length; j++)
            {
                Assert.AreEqual(array2[j],
                             output[j], "Checking property offset " + j);
            }
        }

        private OPOIFSDocument makeCopy(OPOIFSDocument document, byte[] input,
                                       byte[] data)
        {
            OPOIFSDocument copy = null;

            if (input.Length >= 4096)
            {
                RawDataBlock[] blocks =
                    new RawDataBlock[(input.Length + 511) / 512];
                MemoryStream stream = new MemoryStream(data);
                int index = 0;

                while (true)
                {
                    RawDataBlock block = new RawDataBlock(stream);

                    if (block.EOF)
                    {
                        break;
                    }
                    blocks[index++] = block;
                }
                copy = new OPOIFSDocument("test" + input.Length, blocks,
                                         input.Length);
            }
            else
            {
                copy = new OPOIFSDocument(
                    "test" + input.Length,
                    (SmallDocumentBlock[])document.SmallBlocks,
                    input.Length);
            }
            return copy;
        }

        private void checkDocument(OPOIFSDocument document,
                                   byte[] input)
        {
            int big_blocks = 0;
            int small_blocks = 0;
            int total_output = 0;

            if (input.Length >= 4096)
            {
                big_blocks = (input.Length + 511) / 512;
                total_output = big_blocks * 512;
            }
            else
            {
                small_blocks = (input.Length + 63) / 64;
                total_output = 0;
            }
            checkValues(
                big_blocks, small_blocks, total_output,
                makeCopy(
                document, input,
                checkValues(
                    big_blocks, small_blocks, total_output, document,
                    input)), input);
        }

        private byte[] checkValues(int big_blocks, int small_blocks,
                                    int total_output, OPOIFSDocument document,
                                    byte[] input)
        {
            Assert.AreEqual(document, document.DocumentProperty.Document);
            int increment = (int)Math.Sqrt(input.Length);

            for (int j = 1; j <= input.Length; j += increment)
            {
                byte[] buffer = new byte[j];
                int offset = 0;

                for (int k = 0; k < (input.Length / j); k++)
                {
                    document.Read(buffer, offset);
                    for (int n = 0; n < buffer.Length; n++)
                    {
                        Assert.AreEqual(input[(k * j) + n], buffer[n]
                            , "checking byte " + (k * j) + n);
                    }
                    offset += j;
                }
            }
            Assert.AreEqual(big_blocks, document.CountBlocks);
            Assert.AreEqual(small_blocks, document.SmallBlocks.Length);
            MemoryStream stream = new MemoryStream();

            document.WriteBlocks(stream);
            byte[] output = stream.ToArray();

            Assert.AreEqual(total_output, output.Length);
            int limit = Math.Min(total_output, input.Length);

            for (int j = 0; j < limit; j++)
            {
                Assert.AreEqual(input[j],
                             output[j], "Checking document offset " + j);
            }
            for (int j = limit; j < output.Length; j++)
            {
                Assert.AreEqual(unchecked((byte)-1),
                             output[j], "Checking document offset " + j);
            }
            return output;
        }

    }
}