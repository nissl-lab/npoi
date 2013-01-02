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
    /**
     * Class to Test HeaderBlockReader functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestHeaderBlockReading
    {

        /**
         * Constructor TestHeaderBlockReader
         *
         * @param name
         */

        public TestHeaderBlockReading()
        {
        }

        /**
         * Test creating a HeaderBlockReader
         *
         * @exception IOException
         */
        [Test]
        public void TestConstructors()
        {
            string[] hexData = 
        {
                "D0 CF 11 E0 A1 B1 1A E1 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 3B 00 03 00 FE FF 09 00",
                "06 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 FE FF FF FF 00 00 00 00 00 10 00 00 FE FF FF FF",
                "01 00 00 00 FE FF FF FF 00 00 00 00 FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF",
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

            byte[] content = RawDataUtil.Decode(hexData);
            HeaderBlockReader block = new HeaderBlockReader(new MemoryStream(content));

            Assert.AreEqual(-2, block.PropertyStart);

            // verify we can't Read a short block
            byte[] shortblock = new byte[511];

            Array.Copy(content, 0, shortblock, 0, 511);
            try
            {
                block = new HeaderBlockReader(new MemoryStream(shortblock));
                Assert.Fail("Should have caught IOException Reading a short block");
            }
            catch (IOException)
            {

                // as expected
            }

            // try various forms of corruption
            for (int index = 0; index < 8; index++)
            {
                content[index] = (byte)(content[index] - 1);
                try
                {
                    block = new HeaderBlockReader(new MemoryStream(content));
                    Assert.Fail("Should have caught IOException corrupting byte " + index);
                }
                catch (IOException)
                {

                    // as expected
                }

                // restore byte value
                content[index] = (byte)(content[index] + 1);
            }
        }
    }
}