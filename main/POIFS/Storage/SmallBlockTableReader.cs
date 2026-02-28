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

using NPOI.POIFS.Properties;
using NPOI.POIFS.Common;

namespace NPOI.POIFS.Storage
{
    /// <summary>
    /// This class implements reading the small document block list from an
    /// existing file
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class SmallBlockTableReader
    {
        private static BlockList prepareSmallDocumentBlocks(
            POIFSBigBlockSize bigBlockSize,
            RawDataBlockList blockList, RootProperty root,
            int sbatStart)
        {
            // Fetch the blocks which hold the Small Blocks stream
            ListManagedBlock[]
            smallBlockBlocks =
                blockList.FetchBlocks(root.StartBlock, -1);

            // Turn that into a list
            BlockList list = new SmallDocumentBlockList(
                    SmallDocumentBlock.Extract(bigBlockSize, smallBlockBlocks));

            return list;
        }
        private static BlockAllocationTableReader prepareReader(
                POIFSBigBlockSize bigBlockSize,
                RawDataBlockList blockList, BlockList list,
                RootProperty root, int sbatStart)
        {
            // Process the SBAT and blocks
            return new BlockAllocationTableReader(bigBlockSize,
                    blockList.FetchBlocks(sbatStart, -1),
                    list);
        }

        /// <summary>
        /// fetch the small document block list from an existing file, normally
        /// needed for debugging and low level dumping. You should typically call
        /// </summary>
        /// <param name="bigBlockSize">the poifs bigBlockSize</param>
        /// <param name="blockList">the raw data from which the small block table will be extracted</param>
        /// <param name="root">the root property (which contains the start block and small block table size)</param>
        /// <param name="sbatStart">the start block of the SBAT</param>
        /// <returns>the small document block reader</returns>
        public static BlockAllocationTableReader _getSmallDocumentBlockReader(
             POIFSBigBlockSize bigBlockSize,
             RawDataBlockList blockList, RootProperty root,
             int sbatStart)
        {
            BlockList list = prepareSmallDocumentBlocks(
                     bigBlockSize, blockList, root, sbatStart);
            return prepareReader(
                    bigBlockSize, blockList, list, root, sbatStart);
        }

        /// <summary>
        /// fetch the small document block list from an existing file
        /// </summary>
        /// <param name="bigBlockSize">the poifs bigBlockSize</param>
        /// <param name="blockList">the raw data from which the small block table will be extracted</param>
        /// <param name="root">the root property (which contains the start block and small block table size)</param>
        /// <param name="sbatStart">the start block of the SBAT</param>
        /// <returns>the small document block list</returns>
        public static BlockList GetSmallDocumentBlocks(POIFSBigBlockSize bigBlockSize,
                            RawDataBlockList blockList, RootProperty root,
                            int sbatStart)
        {
            BlockList list =
                new SmallDocumentBlockList(
                    SmallDocumentBlock.Extract(bigBlockSize, blockList.FetchBlocks(root.StartBlock, -1)));

            new BlockAllocationTableReader(bigBlockSize, blockList.FetchBlocks(sbatStart, -1), list);
            return list;
        }
    }
}