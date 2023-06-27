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
using NPOI.POIFS.Common;
using NPOI.Util;
using NPOI.POIFS.FileSystem;
using TestCases.Util;

namespace TestCases.POIFS.Storage
{
    /**
     * Class to Test RawDataBlockList functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestRawDataBlockList
    {

        /**
         * Constructor TestRawDataBlockList
         *
         * @param name
         */
        public TestRawDataBlockList()
        {

        }

        /**
         * Test creating a normal RawDataBlockList
         *
         * @exception IOException
         */
        [Test]
        public void TestNormalConstructor()
        {
            byte[] data = new byte[2560];

            for (int j = 0; j < 2560; j++)
            {
                data[j] = (byte)j;
            }
            new RawDataBlockList(new MemoryStream(data), POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
        }

        /**
         * Test creating an empty RawDataBlockList
         *
         * @exception IOException
         */
        [Test]
        public void TestEmptyConstructor()
        {
            new RawDataBlockList(new MemoryStream(new byte[0]), POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
        }

        /**
         * Test creating a short RawDataBlockList
         */
        [Test]
        public void TestShortConstructor()
        {
            // Get the logger to be used
            POILogger logger = POILogFactory.GetLogger(
                    typeof(RawDataBlock)
            );
            if (!(logger is DummyPOILogger dummyPoiLogger))
            {
                // NET Core
                Assert.Ignore("Logger configuration not working under NET Core");
                return;
            }

            dummyPoiLogger.Reset(); // the logger may have been used before
            Assert.AreEqual(0, dummyPoiLogger.logged.Count);

            // Test for various short sizes
            for (int k = 2049; k < 2560; k++)
            {
                byte[] data = new byte[k];

                for (int j = 0; j < k; j++)
                {
                    data[j] = (byte)j;
                }

                // Check we logged the error
                dummyPoiLogger.Reset();
                new RawDataBlockList(new MemoryStream(data), POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
                Assert.AreEqual(1, dummyPoiLogger.logged.Count);
            }
        }
    }
}