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

using System.Collections.Generic;
using System.IO;

using NPOI.POIFS.Common;
using NPOI.POIFS.FileSystem;
using NPOI.Util;

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
    /// *
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class BlockAllocationTableWriter : BlockWritable, BATManaged
    {
        private List<int>    _entries;
        private BATBlock[] _blocks;
        private int        _start_block;
        private POIFSBigBlockSize _bigBlockSize;

        private static int _default_size = 128;

        /// <summary>
        /// Initializes a new instance of the <see cref="BlockAllocationTableWriter"/> class.
        /// </summary>
        public BlockAllocationTableWriter(POIFSBigBlockSize bigBlockSize)
        {
            _start_block = POIFSConstants.END_OF_CHAIN;
            _entries = new List<int>(_default_size);
            _blocks      = new BATBlock[ 0 ];
            _bigBlockSize = bigBlockSize;
        }

        /// <summary>
        /// Create the BATBlocks we need
        /// </summary>
        /// <returns>start block index of BAT blocks</returns>
        public int CreateBlocks()
        {
            int xbat_blocks = 0;
            int bat_blocks  = 0;

            while (true)
            {
                int calculated_bat_blocks  =
                    BATBlock.CalculateStorageRequirements(_bigBlockSize,
                                                          bat_blocks
                                                          + xbat_blocks
                                                          + _entries.Count);
                int calculated_xbat_blocks =
                    HeaderBlockWriter.CalculateXBATStorageRequirements(_bigBlockSize,calculated_bat_blocks);

                if ((bat_blocks == calculated_bat_blocks)
                        && (xbat_blocks == calculated_xbat_blocks))
                {

                    // stable ... we're OK
                    break;
                }
                else
                {
                    bat_blocks  = calculated_bat_blocks;
                    xbat_blocks = calculated_xbat_blocks;
                }
            }
            int startBlock = AllocateSpace(bat_blocks);

            AllocateSpace(xbat_blocks);
            SimpleCreateBlocks();
            return startBlock;
        }

        /// <summary>
        /// Allocate space for a block of indices
        /// </summary>
        /// <param name="blockCount">the number of blocks to allocate space for</param>
        /// <returns>the starting index of the blocks</returns>
        public int AllocateSpace(int blockCount)
        {
            int startBlock = _entries.Count;

            if (blockCount > 0)
            {
                int limit = blockCount - 1;
                int index = startBlock + 1;

                for (int k = 0; k < limit; k++)
                {
                    _entries.Add(index++);
                }
                _entries.Add(POIFSConstants.END_OF_CHAIN);
            }
            return startBlock;
        }

        /// <summary>
        /// Sets the start block for this instance
        /// </summary>
        /// <value>
        /// index into the array of BigBlock instances making up the the filesystem
        /// </value>
        public int StartBlock
        {
            get { return _start_block; }
            set { _start_block = value; }
        }

        /// <summary>
        /// create the BATBlocks
        /// </summary>
        internal void SimpleCreateBlocks()
        {
            _blocks = BATBlock.CreateBATBlocks(_bigBlockSize, _entries.ToArray());
        }

        /// <summary>
        /// Write the storage to an OutputStream
        /// </summary>
        /// <param name="stream">the OutputStream to which the stored data should be written</param>
        public void WriteBlocks(Stream stream)
        {
            for (int j = 0; j < _blocks.Length; j++)
            {
                _blocks[ j ].WriteBlocks(stream);
            }
        }

        //public static void WriteBlock(BATBlock bat, byte[] block)
        public static void WriteBlock(BATBlock bat, ByteBuffer block)
        {
            bat.WriteData(block);
        }

        /// <summary>
        /// Gets the number of BigBlock's this instance uses
        /// </summary>
        /// <value>count of BigBlock instances</value>
        public int CountBlocks
        {
            get { return _blocks.Length; }
        }
    }
}
