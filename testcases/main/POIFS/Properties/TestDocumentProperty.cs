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
using System.Collections;
using System.IO;

using NUnit.Framework;

using NPOI.POIFS.Common;
using NPOI.POIFS.Storage;
using NPOI.POIFS.Properties;
using TestCases.POIFS.Storage;

namespace TestCases.POIFS.Properties
{
    /**
     * Class to Test DocumentProperty functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestDocumentProperty
    {

        /**
         * Constructor TestDocumentProperty
         *
         * @param name
         */

        public TestDocumentProperty()
        {

        }

        /**
         * Test constructing DocumentPropertys
         *
         * @exception IOException
         */
        [Test]
        public void TestConstructor()
        {

            // Test with short name, small file
            VerifyProperty("foo", 1234);

            // Test with just long enough name, small file
            VerifyProperty("A.really.long.long.long.name123", 2345);

            // Test with longer name, just small enough file
            VerifyProperty("A.really.long.long.long.name1234", 4095);

            // Test with just long enough file
            VerifyProperty("A.really.long.long.long.name123", 4096);
        }

        /**
         * Test Reading constructor
         *
         * @exception IOException
         */
        [Test]
        public void TestReadingConstructor()
        {
            string[] hexData = {
            "52 00 6F 00 6F 00 74 00 20 00 45 00 6E 00 74 00 72 00 79 00 00 00 00 00 00 00 00 00 00 00 00 00",
            "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00",
            "16 00 05 01 FF FF FF FF FF FF FF FF 02 00 00 00 20 08 02 00 00 00 00 00 C0 00 00 00 00 00 00 46",
            "00 00 00 00 00 00 00 00 00 00 00 00 C0 5C E8 23 9E 6B C1 01 FE FF FF FF 00 00 00 00 00 00 00 00",
            "57 00 6F 00 72 00 6B 00 62 00 6F 00 6F 00 6B 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00",
            "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00",
            "12 00 02 01 FF FF FF FF FF FF FF FF FF FF FF FF 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00",
            "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 10 00 00 00 00 00 00",
            "05 00 53 00 75 00 6D 00 6D 00 61 00 72 00 79 00 49 00 6E 00 66 00 6F 00 72 00 6D 00 61 00 74 00", //SummaryInformation
            "69 00 6F 00 6E 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00",
            "28 00 02 01 01 00 00 00 03 00 00 00 FF FF FF FF 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00",
            "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 08 00 00 00 00 10 00 00 00 00 00 00",
            "05 00 44 00 6F 00 63 00 75 00 6D 00 65 00 6E 00 74 00 53 00 75 00 6D 00 6D 00 61 00 72 00 79 00",  //DocumentSummaryInformation
            "49 00 6E 00 66 00 6F 00 72 00 6D 00 61 00 74 00 69 00 6F 00 6E 00 00 00 00 00 00 00 00 00 00 00",
            "38 00 02 01 FF FF FF FF FF FF FF FF FF FF FF FF 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00",
            "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 10 00 00 00 00 10 00 00 00 00 00 00",
        };
            byte[] input = RawDataUtil.Decode(hexData);

            VerifyReadingProperty(1, input, 128, "Workbook");
            VerifyReadingProperty(2, input, 256, "\x0005SummaryInformation");
            VerifyReadingProperty(3, input, 384, "\x0005DocumentSummaryInformation");
        }

        private void VerifyReadingProperty(int index, byte[] input, int offset, string name)
        {
            DocumentProperty property = new DocumentProperty(index, input, offset);
            MemoryStream stream = new MemoryStream(128);
            byte[] expected = new byte[128];

            Array.Copy(input, offset, expected, 0, 128);
            property.WriteData(stream);
            byte[] output = stream.ToArray();

            Assert.AreEqual(128, output.Length);
            for (int j = 0; j < 128; j++)
            {
                Assert.AreEqual(expected[j],
                             output[j], "mismatch at offset " + j);
            }
            Assert.AreEqual(index, property.Index);
            Assert.AreEqual(name, property.Name);
        }

        private void VerifyProperty(String name, int size)
        {
            DocumentProperty property = new DocumentProperty(name, size);

            if (size >= 4096)
            {
                Assert.IsTrue(!property.ShouldUseSmallBlocks);
            }
            else
            {
                Assert.IsTrue(property.ShouldUseSmallBlocks);
            }
            byte[] Testblock = new byte[128];
            int index = 0;

            for (; index < 0x40; index++)
            {
                Testblock[index] = (byte)0;
            }
            int limit = Math.Min(31, name.Length);

            Testblock[index++] = (byte)(2 * (limit + 1));
            Testblock[index++] = (byte)0;
            Testblock[index++] = (byte)2;
            Testblock[index++] = (byte)1;
            for (; index < 0x50; index++)
            {
                Testblock[index] = (byte)0xFF;
            }
            for (; index < 0x78; index++)
            {
                Testblock[index] = (byte)0;
            }
            int sz = size;

            Testblock[index++] = (byte)sz;
            sz /= 256;
            Testblock[index++] = (byte)sz;
            sz /= 256;
            Testblock[index++] = (byte)sz;
            sz /= 256;
            Testblock[index++] = (byte)sz;
            for (; index < 0x80; index++)
            {
                Testblock[index] = (byte)0x0;
            }
            byte[] name_bytes = Encoding.GetEncoding(1252).GetBytes(name);

            for (index = 0; index < limit; index++)
            {
                Testblock[index * 2] = name_bytes[index];
            }
            MemoryStream stream = new MemoryStream(512);

            property.WriteData(stream);
            byte[] output = stream.ToArray();

            Assert.AreEqual(Testblock.Length, output.Length);
            for (int j = 0; j < Testblock.Length; j++)
            {
                Assert.AreEqual(Testblock[j],
                             output[j], "mismatch at offset " + j);
            }
        }
    }
}