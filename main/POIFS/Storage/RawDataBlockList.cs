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

using System.IO;
using NPOI.POIFS.Common;
using System.Collections.Generic;

namespace NPOI.POIFS.Storage
{
    /// <summary>
    /// A list of RawDataBlocks instances, and methods to manage the list
    /// @author Marc Johnson (mjohnson at apache dot org
    /// </summary>
    public class RawDataBlockList:BlockListImpl
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="RawDataBlockList"/> class.
        /// </summary>
        /// <param name="stream">the InputStream from which the data will be read</param>
        /// <param name="bigBlockSize">The big block size, either 512 bytes or 4096 bytes</param>
        public RawDataBlockList(Stream stream, POIFSBigBlockSize bigBlockSize)
        {
            List<RawDataBlock> blocks = new List<RawDataBlock>();

            while (true)
            {
                RawDataBlock block = new RawDataBlock(stream, bigBlockSize.GetBigBlockSize());
                
                // If there was data, add the block to the list
                if(block.HasData) {
            	    blocks.Add(block);
                }

                // If the stream is now at the End Of File, we're done
                if (block.EOF) {
                    break;
                }
            }
             SetBlocks((ListManagedBlock[])blocks.ToArray());
        }
    }
}
