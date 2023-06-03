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
using TestCases.Util;
using TestCases.POIFS.FileSystem;

namespace TestCases.POIFS.Storage
{
    /**
     * Class to Test RawDataBlock functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestRawDataBlock
    {
        public TestRawDataBlock()
        {

        }

        /**
         * Test creating a normal RawDataBlock
         *
         * @exception IOException
         */
        [Test]
        public void TestNormalConstructor()
        {
            byte[] data = new byte[512];

            for (int j = 0; j < 512; j++)
            {
                data[j] = (byte)j;
            }
            RawDataBlock block = new RawDataBlock(new MemoryStream(data));

            Assert.IsTrue(!block.EOF, "Should not be at EOF");
            byte[] out_data = block.Data;

            Assert.AreEqual(data.Length, out_data.Length, "Should be same Length");
            for (int j = 0; j < 512; j++)
            {
                Assert.AreEqual(data[j],
                             out_data[j], "Should be same value at offset " + j);
            }
        }

        /**
         * Test creating an empty RawDataBlock
         *
         * @exception IOException
         */
        [Test]
        public void TestEmptyConstructor()
        {
            byte[] data = new byte[0];
            RawDataBlock block = new RawDataBlock(new MemoryStream(data));

            Assert.IsTrue(block.EOF, "Should be at EOF");
            try
            {
                byte[] a = block.Data;
            }
            catch (IOException )
            {

                // as expected
            }
        }

        /**
         * Test creating a short RawDataBlock
         * Will trigger a warning, but no longer an IOException,
         *  as people seem to have "valid" truncated files
         */
        [Test]
        public void TestShortConstructor()
        {
            //// Get the logger to be used
            POILogger logger = POILogFactory.GetLogger(typeof(RawDataBlock));
            if (!(logger is DummyPOILogger dummyPoiLogger))
            {
                // NET Core
                Assert.Ignore("Logger configuration not working under NET Core");
                return;
            }

            dummyPoiLogger.Reset(); // the logger may have been used before
            Assert.AreEqual(0, dummyPoiLogger.logged.Count);

            // Test for various data sizes
            for (int k = 1; k <= 512; k++)
            {
                byte[] data = new byte[k];

                for (int j = 0; j < k; j++)
                {
                    data[j] = (byte)j;
                }
                RawDataBlock block = null;

                dummyPoiLogger.Reset();
                Assert.AreEqual(0, dummyPoiLogger.logged.Count);

                // Have it created
                block = new RawDataBlock(new MemoryStream(data));
                Assert.IsNotNull(block);

                // Check for the warning Is there for <512
                if (k < 512)
                {
                    Assert.AreEqual(
                            1, dummyPoiLogger.logged.Count, "Warning on " + k + " byte short block"
                    );

                    // Build the expected warning message, and check
                    String bts = k + " byte";
                    if (k > 1)
                    {
                        bts += "s";
                    }

                    Assert.AreEqual(
                            (String)dummyPoiLogger.logged[0],
                            "7 - Unable to read entire block; " + bts + " read before EOF; expected 512 bytes. Your document was either written by software that ignores the spec, or has been truncated!"
                    );
                }
                else
                {
                    Assert.AreEqual(0, dummyPoiLogger.logged.Count);
                }
            }
        }

        /**
         * Tests that when using a slow input stream, which
         *  won't return a full block at a time, we don't
         *  incorrectly think that there's not enough data
         */
        [Test]
        public void TestSlowInputStream()
        {
            // Get the logger to be used
            POILogger logger = POILogFactory.GetLogger(typeof(RawDataBlock));

            if (!(logger is DummyPOILogger dummyPoiLogger))
            {
                // NET Core
                Assert.Ignore("Logger configuration not working under NET Core");
                return;
            }

            dummyPoiLogger.Reset(); // the logger may have been used before
            Assert.AreEqual(0, dummyPoiLogger.logged.Count);

            // Test for various ok data sizes
            for (int k = 1; k < 512; k++)
            {
                byte[] data = new byte[512];
                for (int j = 0; j < data.Length; j++)
                {
                    data[j] = (byte)j;
                }

                // Shouldn't complain, as there Is enough data,
                //  even if it dribbles through
                RawDataBlock block =
                    new RawDataBlock(new SlowInputStream(data, 512));   //k is changed to 512
                Assert.IsFalse(block.EOF);
            }

            // But if there wasn't enough data available, will
            //  complain
            for (int k = 1; k < 512; k++)
            {
                byte[] data = new byte[511];
                for (int j = 0; j < data.Length; j++)
                {
                    data[j] = (byte)j;
                }

                dummyPoiLogger.Reset();
                Assert.AreEqual(0, dummyPoiLogger.logged.Count);

                // Should complain, as there Isn't enough data
                RawDataBlock block =
                    new RawDataBlock(new SlowInputStream(data, k));
                Assert.IsNotNull(block);
                Assert.AreEqual(
                        1, dummyPoiLogger.logged.Count, "Warning on " + k + " byte short block"
                );
            }
        }


    }
}