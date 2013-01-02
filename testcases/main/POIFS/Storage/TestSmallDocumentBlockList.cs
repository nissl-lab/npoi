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
     * Class to Test SmallDocumentBlockList functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestSmallDocumentBlockList
    {
        /**
         * Test creating a SmallDocumentBlockList
         *
         * @exception IOException
         */
        [Test]
        public void TestConstructor()
        {
            byte[] data = new byte[2560];

            for (int j = 0; j < 2560; j++)
            {
                data[j] = (byte)j;
            }
            MemoryStream stream = new MemoryStream(data);
            RawDataBlock[] blocks = new RawDataBlock[5];

            for (int j = 0; j < 5; j++)
            {
                blocks[j] = new RawDataBlock(stream);
            }
            SmallDocumentBlockList sdbl =
                new SmallDocumentBlockList(SmallDocumentBlock.Extract(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, blocks));

            // proof we added the blocks
            for (int j = 0; j < 40; j++)
            {
                sdbl.Remove(j);
            }
            try
            {
                sdbl.Remove(41);
                Assert.Fail("there should have been an Earth-shattering ka-boom!");
            }
            catch (IOException )
            {

                // it better have thrown one!!
            }
        }


    }
}