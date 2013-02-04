
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.POIFS.FileSystem
{
    using System;
    using NPOI.POIFS.Storage;
    using NPOI.Util;

    /// <summary>
    /// This abstract class describes a way to read, store, chain
    /// and free a series of blocks (be they Big or Small ones)
    /// </summary>
    public abstract class BlockStore
    {
        /// <summary>
        /// Returns the size of the blocks managed through the block store.
        /// </summary>
        /// <returns></returns>
        public abstract int GetBlockStoreBlockSize();


        /// <summary>
        /// Load the block at the given offset.
        /// </summary>
        /// <param name="offset"></param>
        /// <returns></returns>
        public abstract ByteBuffer GetBlockAt(int offset);
        /// <summary>
        /// Extends the file if required to hold blocks up to
        /// the specified offset, and return the block from there.
        /// </summary>
        /// <param name="offset"></param>
        /// <returns></returns>
        public abstract ByteBuffer CreateBlockIfNeeded(int offset);
        /// <summary>
        /// Returns the BATBlock that handles the specified offset,
        /// and the relative index within it
        /// </summary>
        /// <param name="offset"></param>
        /// <returns></returns>
        public abstract BATBlockAndIndex GetBATBlockAndIndex(int offset);

        /// <summary>
        /// Works out what block follows the specified one.
        /// </summary>
        /// <param name="offset"></param>
        /// <returns></returns>
        public abstract int GetNextBlock(int offset);

        /// <summary>
        /// Changes the record of what block follows the specified one.
        /// </summary>
        /// <param name="offset"></param>
        /// <param name="nextBlock"></param>
        public abstract void SetNextBlock(int offset, int nextBlock);
        /// <summary>
        /// Finds a free block, and returns its offset.
        /// This method will extend the file/stream if needed, and if doing
        ///  so, allocate new FAT blocks to address the extra space.
        /// </summary>
        /// <returns></returns>
        public abstract int GetFreeBlock();
        /// <summary>
        /// Creates a Detector for loops in the chain 
        /// </summary>
        /// <returns></returns>
        public abstract ChainLoopDetector GetChainLoopDetector();

        
    }
        /// <summary>
        /// Used to detect if a chain has a loop in it, so
        ///  we can bail out with an error rather than
        ///  spinning away for ever... 
        /// </summary>
    public class ChainLoopDetector
        {
            private bool[] used_blocks;
        private BlockStore blockStore;

        public ChainLoopDetector(long rawSize, BlockStore blockStore)
            {
            this.blockStore = blockStore;
            int numBlocks = (int)Math.Ceiling(1.0 * (rawSize / blockStore.GetBlockStoreBlockSize()));
                used_blocks = new bool[numBlocks];
            }

        public void Claim(int offset)
            {
                if (offset >= used_blocks.Length)
                {
                    // They're writing, and have had new blocks requested
                    //  for the write to proceed. That means they're into
                    //  blocks we've allocated for them, so are safe
                    return;
                }

                // Claiming an existing block, ensure there's no loop
                if (used_blocks[offset])
                {
                    throw new InvalidOperationException(
                          "Potential loop detected - Block " + offset +
                          " was already claimed but was just requested again"
                    );
                }
                used_blocks[offset] = true;
            }
        }

}