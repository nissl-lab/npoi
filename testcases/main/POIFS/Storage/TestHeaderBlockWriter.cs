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

    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.POIFS.Storage;
    using NPOI.Util;
    using NPOI.POIFS.FileSystem;
    /**
     * Class to Test HeaderBlockWriter functionality
     *
     * @author Marc Johnson
     */
    [TestClass]
    public class TestHeaderBlockWriter
    {

        /**
         * Constructor TestHeaderBlockWriter
         *
         * @param name
         */

        public TestHeaderBlockWriter()
        {
        }

        /**
         * Test creating a HeaderBlockWriter
         *
         * @exception IOException
         */
        [TestMethod]
        public void TestConstructors()
        {
            HeaderBlockWriter block = new HeaderBlockWriter();
            MemoryStream output = new MemoryStream(512);

            block.WriteBlocks(output);
            byte[] copy = output.ToArray();
            byte[] expected =
        {
            ( byte ) 0xD0, ( byte ) 0xCF, ( byte ) 0x11, ( byte ) 0xE0,
            ( byte ) 0xA1, ( byte ) 0xB1, ( byte ) 0x1A, ( byte ) 0xE1,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x3B, ( byte ) 0x00, ( byte ) 0x03, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0x09, ( byte ) 0x00,
            ( byte ) 0x06, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x10, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF
        };

            Assert.AreEqual(expected.Length, copy.Length);
            for (int j = 0; j < 512; j++)
            {
                Assert.AreEqual(expected[j], copy[j], "testing byte " + j);
            }

            // verify we can Read a 'good' HeaderBlockWriter (also Test
            // GetPropertyStart)
            block.PropertyStart = unchecked((int)0x87654321);
            output = new MemoryStream(512);
            block.WriteBlocks(output);
            Assert.AreEqual(unchecked((int)0x87654321),
                         new HeaderBlockReader(new MemoryStream(output
                             .ToArray())).PropertyStart);
        }

        /**
         * Test Setting the SBAT start block
         *
         * @exception IOException
         */
        [TestMethod]
        public void TestSetSBATStart()
        {
            HeaderBlockWriter block = new HeaderBlockWriter();

            block.SBATStart = 0x01234567;
            MemoryStream output = new MemoryStream(512);

            block.WriteBlocks(output);
            byte[] copy = output.ToArray();
            byte[] expected =
        {
            ( byte ) 0xD0, ( byte ) 0xCF, ( byte ) 0x11, ( byte ) 0xE0,
            ( byte ) 0xA1, ( byte ) 0xB1, ( byte ) 0x1A, ( byte ) 0xE1,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x3B, ( byte ) 0x00, ( byte ) 0x03, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0x09, ( byte ) 0x00,
            ( byte ) 0x06, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x10, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x67, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF
        };

            Assert.AreEqual(expected.Length, copy.Length);
            for (int j = 0; j < 512; j++)
            {
                Assert.AreEqual(expected[j], copy[j], "testing byte " + j);
            }
        }

        /**
         * Test SetPropertyStart and GetPropertyStart
         *
         * @exception IOException
         */
        [TestMethod]
        public void TestSetPropertyStart()
        {
            HeaderBlockWriter block = new HeaderBlockWriter();

            block.PropertyStart = 0x01234567;
            MemoryStream output = new MemoryStream(512);

            block.WriteBlocks(output);
            byte[] copy = output.ToArray();
            byte[] expected =
        {
            ( byte ) 0xD0, ( byte ) 0xCF, ( byte ) 0x11, ( byte ) 0xE0,
            ( byte ) 0xA1, ( byte ) 0xB1, ( byte ) 0x1A, ( byte ) 0xE1,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x3B, ( byte ) 0x00, ( byte ) 0x03, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0x09, ( byte ) 0x00,
            ( byte ) 0x06, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x67, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x10, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF
        };

            Assert.AreEqual(expected.Length, copy.Length);
            for (int j = 0; j < 512; j++)
            {
                Assert.AreEqual(expected[j], copy[j], "testing byte " + j);
            }
        }

        /**
         * Test Setting the BAT blocks; also Tests GetBATCount,
         * GetBATArray, GetXBATCount
         *
         * @exception IOException
         */
        [TestMethod]
        public void TestSetBATBlocks()
        {

            // first, a small Set of blocks
            HeaderBlockWriter block = new HeaderBlockWriter();
            BATBlock[] xbats = block.SetBATBlocks(5, 0x01234567);

            Assert.AreEqual(0, xbats.Length);
            Assert.AreEqual(0, HeaderBlockWriter
                .CalculateXBATStorageRequirements(5));
            MemoryStream output = new MemoryStream(512);

            block.WriteBlocks(output);
            byte[] copy = output.ToArray();
            byte[] expected =
        {
            ( byte ) 0xD0, ( byte ) 0xCF, ( byte ) 0x11, ( byte ) 0xE0,
            ( byte ) 0xA1, ( byte ) 0xB1, ( byte ) 0x1A, ( byte ) 0xE1,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x3B, ( byte ) 0x00, ( byte ) 0x03, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0x09, ( byte ) 0x00,
            ( byte ) 0x06, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x05, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x10, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x67, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x68, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x69, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x6A, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x6B, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF
        };

            Assert.AreEqual(expected.Length, copy.Length);
            for (int j = 0; j < 512; j++)
            {
                Assert.AreEqual(expected[j], copy[j], "testing byte " + j);
            }

            // second, a full Set of blocks (109 blocks)
            block = new HeaderBlockWriter();
            xbats = block.SetBATBlocks(109, 0x01234567);
            Assert.AreEqual(0, xbats.Length);
            Assert.AreEqual(0, HeaderBlockWriter
                .CalculateXBATStorageRequirements(109));
            output = new MemoryStream(512);
            block.WriteBlocks(output);
            copy = output.ToArray();
            byte[] expected2 =
        {
            ( byte ) 0xD0, ( byte ) 0xCF, ( byte ) 0x11, ( byte ) 0xE0,
            ( byte ) 0xA1, ( byte ) 0xB1, ( byte ) 0x1A, ( byte ) 0xE1,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x3B, ( byte ) 0x00, ( byte ) 0x03, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0x09, ( byte ) 0x00,
            ( byte ) 0x06, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x6D, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x10, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x67, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x68, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x69, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x6A, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x6B, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x6C, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x6D, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x6E, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x6F, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x70, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x71, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x72, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x73, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x74, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x75, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x76, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x77, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x78, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x79, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x7A, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x7B, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x7C, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x7D, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x7E, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x7F, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x80, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x81, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x82, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x83, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x84, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x85, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x86, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x87, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x88, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x89, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x8A, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x8B, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x8C, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x8D, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x8E, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x8F, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x90, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x91, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x92, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x93, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x94, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x95, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x96, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x97, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x98, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x99, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x9A, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x9B, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x9C, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x9D, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x9E, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x9F, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA0, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA1, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA2, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA3, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA4, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA5, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA6, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA7, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA8, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA9, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xAA, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xAB, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xAC, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xAD, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xAE, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xAF, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB0, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB1, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB2, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB3, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB4, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB5, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB6, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB7, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB8, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB9, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xBA, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xBB, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xBC, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xBD, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xBE, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xBF, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC0, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC1, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC2, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC3, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC4, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC5, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC6, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC7, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC8, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC9, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xCA, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xCB, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xCC, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xCD, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xCE, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xCF, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xD0, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xD1, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xD2, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xD3, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01
        };

            Assert.AreEqual(expected2.Length, copy.Length);
            for (int j = 0; j < 512; j++)
            {
                Assert.AreEqual(expected2[j], copy[j], "testing byte " + j);
            }

            // finally, a really large Set of blocks (256 blocks)
            block = new HeaderBlockWriter();
            xbats = block.SetBATBlocks(256, 0x01234567);
            Assert.AreEqual(2, xbats.Length);
            Assert.AreEqual(2, HeaderBlockWriter
                .CalculateXBATStorageRequirements(256));
            output = new MemoryStream(512);
            block.WriteBlocks(output);
            copy = output.ToArray();
            byte[] expected3 =
        {
            ( byte ) 0xD0, ( byte ) 0xCF, ( byte ) 0x11, ( byte ) 0xE0,
            ( byte ) 0xA1, ( byte ) 0xB1, ( byte ) 0x1A, ( byte ) 0xE1,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x3B, ( byte ) 0x00, ( byte ) 0x03, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0x09, ( byte ) 0x00,
            ( byte ) 0x06, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x01, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x10, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x67, ( byte ) 0x46, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x02, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x67, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x68, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x69, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x6A, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x6B, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x6C, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x6D, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x6E, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x6F, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x70, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x71, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x72, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x73, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x74, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x75, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x76, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x77, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x78, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x79, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x7A, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x7B, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x7C, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x7D, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x7E, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x7F, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x80, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x81, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x82, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x83, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x84, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x85, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x86, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x87, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x88, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x89, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x8A, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x8B, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x8C, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x8D, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x8E, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x8F, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x90, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x91, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x92, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x93, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x94, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x95, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x96, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x97, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x98, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x99, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x9A, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x9B, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x9C, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x9D, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x9E, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0x9F, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA0, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA1, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA2, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA3, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA4, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA5, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA6, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA7, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA8, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xA9, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xAA, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xAB, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xAC, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xAD, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xAE, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xAF, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB0, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB1, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB2, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB3, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB4, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB5, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB6, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB7, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB8, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xB9, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xBA, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xBB, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xBC, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xBD, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xBE, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xBF, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC0, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC1, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC2, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC3, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC4, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC5, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC6, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC7, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC8, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xC9, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xCA, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xCB, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xCC, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xCD, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xCE, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xCF, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xD0, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xD1, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xD2, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01,
            ( byte ) 0xD3, ( byte ) 0x45, ( byte ) 0x23, ( byte ) 0x01
        };

            Assert.AreEqual(expected3.Length, copy.Length);
            for (int j = 0; j < 512; j++)
            {
                Assert.AreEqual(expected3[j], copy[j], "testing byte " + j);
            }
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