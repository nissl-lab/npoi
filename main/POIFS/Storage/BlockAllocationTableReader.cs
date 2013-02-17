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
using System.Collections.Generic;
using System.IO;

using NPOI.Util;
using NPOI.POIFS.Common;

namespace NPOI.POIFS.Storage
{
    /// <summary>
    /// This class manages and creates the Block Allocation Table, which is
    /// basically a set of linked lists of block indices.
    /// Each block of the filesystem has an index. The first block, the
    /// header, is skipped; the first block after the header is index 0,
    /// the next is index 1, and so on.
    /// A block's index is also its index into the Block Allocation
    /// Table. The entry that it finds in the Block Allocation Table is the
    /// index of the next block in the linked list of blocks making up a
    /// file, or it is set to -2: end of list.
    /// 
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class BlockAllocationTableReader
    {

        private static POILogger _logger = POILogFactory.GetLogger(typeof(BlockAllocationTableReader));

        private const int MAX_BLOCK_COUNT = 65535;
       
        private List<int> _entries;

        private POIFSBigBlockSize bigBlockSize;
        /// <summary>
        /// create a BlockAllocationTableReader for an existing filesystem. Side
        /// effect: when this method finishes, the BAT blocks will have
        /// been Removed from the raw block list, and any blocks labeled as
        /// 'unused' in the block allocation table will also have been
        /// Removed from the raw block list. </summary>
        /// <param name="bigBlockSizse">the poifs bigBlockSize</param>
        /// <param name="block_count">the number of BAT blocks making up the block allocation table</param>
        /// <param name="block_array">the array of BAT block indices from the
        /// filesystem's header</param>
        /// <param name="xbat_count">the number of XBAT blocks</param>
        /// <param name="xbat_index">the index of the first XBAT block</param>
        /// <param name="raw_block_list">the list of RawDataBlocks</param>
        public BlockAllocationTableReader(POIFSBigBlockSize bigBlockSizse,
                                          int block_count,
                                          int[] block_array,
                                          int xbat_count,
                                          int xbat_index,
                                          BlockList raw_block_list)
            : this(bigBlockSizse)
        {
            SanityCheckBlockCount(block_count);

            RawDataBlock[] blocks = new RawDataBlock[block_count];
            int limit = Math.Min(block_count, block_array.Length);
            int block_index;

            for (block_index = 0; block_index < limit; block_index++)
            {
                int nextOffset = block_array[block_index];
                if (nextOffset > raw_block_list.BlockCount())
                {
                    throw new IOException("Your file contains " + raw_block_list.BlockCount() +
                                           " sectors, but the initial DIFAT array at index " + block_index +
                                           " referenced block # " + nextOffset + ". This isn't allowed and " +
                                           " your file is corrupt");
                }

                blocks[block_index] = (RawDataBlock)raw_block_list.Remove(nextOffset);
            }
            if (block_index < block_count)
            {

                // must have extended blocks
                if (xbat_index < 0)
                {
                    throw new IOException(
                        "BAT count exceeds limit, yet XBAT index indicates no valid entries");
                }
                int chain_index = xbat_index;
                int max_entries_per_block = BATBlock.EntriesPerXBATBlock;
                int chain_index_offset = BATBlock.XBATChainOffset;

                // Each XBAT block contains either:
                //  (maximum number of sector indexes) + index of next XBAT
                //  some sector indexes + FREE sectors to max # + EndOfChain
                for (int j = 0; j < xbat_count; j++)
                {
                    limit = Math.Min(block_count - block_index,
                                     max_entries_per_block);
                    byte[] data = raw_block_list.Remove(chain_index).Data;
                    int offset = 0;

                    for (int k = 0; k < limit; k++)
                    {
                        blocks[block_index++] =
                            (RawDataBlock)raw_block_list.Remove(LittleEndian.GetInt(data, offset));
                        offset += LittleEndianConsts.INT_SIZE;
                    }
                    chain_index = LittleEndian.GetInt(data, chain_index_offset);
                    if (chain_index == POIFSConstants.END_OF_CHAIN)
                    {
                        break;
                    }
                }
            }
            if (block_index != block_count)
            {
                throw new IOException("Could not find all blocks");
            }

            // now that we have all of the raw data blocks, go through and
            // create the indices
            SetEntries((ListManagedBlock[])blocks, raw_block_list);
        }

        /// <summary>
        /// create a BlockAllocationTableReader from an array of raw data blocks
        /// </summary>
        /// <param name="bigBlockSize"></param>
        /// <param name="blocks">the raw data</param>
        /// <param name="raw_block_list">the list holding the managed blocks</param>
        public BlockAllocationTableReader(POIFSBigBlockSize bigBlockSize, ListManagedBlock[] blocks,  
                                   BlockList raw_block_list)
            : this(bigBlockSize)
        {
            SetEntries(blocks, raw_block_list);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="BlockAllocationTableReader"/> class.
        /// </summary>
        public BlockAllocationTableReader(POIFSBigBlockSize bigBlockSize)
        {
            this.bigBlockSize = bigBlockSize;
            _entries = new List<int>();
        }

        /// <summary>
        /// walk the entries from a specified point and return the
        /// associated blocks. The associated blocks are Removed from the block list
        /// </summary>
        /// <param name="startBlock">the first block in the chain</param>
        /// <param name="headerPropertiesStartBlock"></param>
        /// <param name="blockList">the raw data block list</param>
        /// <returns>array of ListManagedBlocks, in their correct order</returns>
        public ListManagedBlock[] FetchBlocks(int startBlock, int headerPropertiesStartBlock,
                                        BlockList blockList)
        {
            List<ListManagedBlock> blocks = new List<ListManagedBlock>();
            int currentBlock = startBlock;
            bool firstPass = true;
            ListManagedBlock dataBlock = null;

            while (currentBlock != POIFSConstants.END_OF_CHAIN)
            {
                try
                {
                    dataBlock = blockList.Remove(currentBlock);
                    blocks.Add(dataBlock);
                    currentBlock = _entries[currentBlock];
                    firstPass = false;
                }
                catch(Exception)
                {
                    if (currentBlock == headerPropertiesStartBlock)
                    {
                        // Special case where things are in the wrong order
                        _logger.Log(POILogger.WARN, "Warning, header block comes after data blocks in POIFS block listing");
                        currentBlock = POIFSConstants.END_OF_CHAIN;
                    }
                    else if (currentBlock == 0 && firstPass)
                    {
                        // Special case where the termination isn't done right
                        //  on an empty set
                        _logger.Log(POILogger.WARN, "Warning, incorrectly terminated empty data blocks in POIFS block listing (should end at -2, ended at 0)");
                        currentBlock = POIFSConstants.END_OF_CHAIN;
                    }
                    else
                    {
                        // Ripple up
                        throw;
                    }
                }
            }
            ListManagedBlock[] array = blocks.ToArray();
            return (array);
        }

        /// <summary>
        /// determine whether the block specified by index is used or not
        /// </summary>
        /// <param name="index">determine whether the block specified by index is used or not</param>
        /// <returns>
        /// 	<c>true</c> if the specified block is used; otherwise, <c>false</c>.
        /// </returns>
        public bool IsUsed(int index)
        {
            bool rval = false;

            try
            {
                rval = _entries[index] != -1;
            }
            catch (IndexOutOfRangeException)
            {
            }
            return rval;
        }

        /// <summary>
        /// return the next block index
        /// </summary>
        /// <param name="index">The index of the current block</param>
        /// <returns>index of the next block (may be
        /// POIFSConstants.END_OF_CHAIN, indicating end of chain
        /// (duh))</returns>
        public int GetNextBlockIndex(int index)
        {
            if (IsUsed(index))
            {
                return _entries[index];
            }
            else
            {
                throw new IOException("index " + index + " is unused");
            }
        }

        /// <summary>
        /// Convert an array of blocks into a Set of integer indices
        /// </summary>
        /// <param name="blocks">the array of blocks containing the indices</param>
        /// <param name="raw_blocks">the list of blocks being managed. Unused
        /// blocks will be eliminated from the list</param>
        private void SetEntries(ListManagedBlock[] blocks,
                                BlockList raw_blocks)
        {

            int limit = bigBlockSize.GetBATEntriesPerBlock();

            for (int block_index = 0; block_index < blocks.Length; block_index++)
            {
                byte[] data = blocks[block_index].Data;
                int offset = 0;

                for (int k = 0; k < limit; k++)
                {
                    int entry = LittleEndian.GetInt(data, offset);

                    if (entry == POIFSConstants.UNUSED_BLOCK)
                    {
                        raw_blocks.Zap(_entries.Count);
                    }
                    _entries.Add(entry);
                    offset += LittleEndianConsts.INT_SIZE;
                }

                // discard block
                blocks[block_index] = null;
            }
            raw_blocks.BAT = this;
        }

        public static void SanityCheckBlockCount(int block_count)
        {
                if (block_count <= 0)
                {
                    throw new IOException("Illegal block count; minimum count is 1, got " 
                                            + block_count + " instead");
                }

                if (block_count > MAX_BLOCK_COUNT)
                {
                    throw new IOException(
                               "Block count " + block_count +
                                " is too high. POI maximum is " + MAX_BLOCK_COUNT + ".");
                }

        }
    }
}