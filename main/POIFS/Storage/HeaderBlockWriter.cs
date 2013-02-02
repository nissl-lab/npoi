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

using NPOI.POIFS.Common;
using NPOI.Util;

namespace NPOI.POIFS.Storage
{
    /// <summary>
    /// The block containing the archive header
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>

    public class HeaderBlockWriter : HeaderBlockConstants, BlockWritable
    {
        private HeaderBlock _header_block;

        public HeaderBlockWriter(POIFSBigBlockSize bigBlockSize)
        {
            _header_block = new HeaderBlock(bigBlockSize);
        }

        public HeaderBlockWriter(HeaderBlock headerBlock)
        {
            _header_block = headerBlock;
        }

        /// <summary>
        /// Set BAT block parameters. Assumes that all BAT blocks are
        /// contiguous. Will construct XBAT blocks if necessary and return
        /// the array of newly constructed XBAT blocks.
        /// </summary>
        /// <param name="blockCount">count of BAT blocks</param>
        /// <param name="startBlock">index of first BAT block</param>
        /// <returns>array of XBAT blocks; may be zero Length, will not be
        /// null</returns>
        public BATBlock[] SetBATBlocks(int blockCount, int startBlock)
        {
            BATBlock[] rvalue;

            POIFSBigBlockSize bigBlockSize = _header_block.BigBlockSize;

            _header_block.BATCount = blockCount;

            int limit = Math.Min(blockCount, _max_bats_in_header);
            int[] bat_blocks = new int[limit];
            for (int j = 0; j < limit; j++)
                bat_blocks[j] = startBlock + j;

            _header_block.BATArray = bat_blocks;

            if (blockCount > _max_bats_in_header)
            {
                int excess_blocks = blockCount - _max_bats_in_header;
                int[] excess_block_array = new int[excess_blocks];

                for (int j = 0; j < excess_blocks; j++)
                    excess_block_array[j] = startBlock + j + _max_bats_in_header;

                rvalue = BATBlock.CreateXBATBlocks(bigBlockSize, excess_block_array,
                                                   startBlock + blockCount);

                _header_block.XBATStart = startBlock + blockCount;
            }
            else
            {
                rvalue = BATBlock.CreateXBATBlocks(bigBlockSize, new int[0], 0);
                _header_block.XBATStart = POIFSConstants.END_OF_CHAIN;
            }

            _header_block.XBATCount = rvalue.Length;
            return rvalue;
        }

        /// <summary>
        /// Set start of Property Table
        /// </summary>
        /// <value>the index of the first block of the Property
        /// Table</value>
        public int PropertyStart
        {
            get { return _header_block.PropertyStart; }
            set { _header_block.PropertyStart = value; }
        }

        /// <summary>
        /// Set start of small block allocation table
        /// </summary>
        /// <value>the index of the first big block of the small
        /// block allocation table</value>
        public int SBAStart
        {
            get { return _header_block.SBATStart; }
            set { _header_block.SBATStart = value; }
        }

        public int SBATStart
        {
            get { return _header_block.SBATStart; }
            set { _header_block.SBATStart = value; }
        }

        /// <summary>
        /// Set count of SBAT blocks
        /// </summary>
        /// <value>the number of SBAT blocks</value>
        public int SBATBlockCount
        {
            get { return _header_block.SBATBlockCount; }
            set { _header_block.SBATBlockCount = value; }
        }

        /// <summary>
        /// For a given number of BAT blocks, calculate how many XBAT
        /// blocks will be needed
        /// </summary>
        /// <param name="bigBlockSize"></param>
        /// <param name="blockCount">number of BAT blocks</param>
        /// <returns>number of XBAT blocks needed</returns>
        public static int CalculateXBATStorageRequirements(POIFSBigBlockSize bigBlockSize, int blockCount)
        {
            return (blockCount > _max_bats_in_header)
                ? BATBlock.CalculateXBATStorageRequirements(bigBlockSize, blockCount - _max_bats_in_header) : 0;
        }


        /// <summary>
        /// Write the block's data to an Stream
        /// </summary>
        /// <param name="stream">the Stream to which the stored data should
        /// be written
        /// </param>
        public void WriteBlocks(Stream stream)
        {
            try
            {
                _header_block.WriteData(stream);
            }
            catch (IOException ex)
            {
                throw ex;
            }
        }

        public void WriteBlock(ByteBuffer block)
        {
            MemoryStream ms = new MemoryStream(_header_block.BigBlockSize.GetBigBlockSize());

            _header_block.WriteData(ms);

            block.Write(ms.ToArray());
        }


        public void WriteBlock(byte[] block)
        {
            MemoryStream ms = new MemoryStream(_header_block.BigBlockSize.GetBigBlockSize());

            _header_block.WriteData(ms);

            //block = ms.ToArray();
            byte[] temp = ms.ToArray();
            Array.Copy(temp, 0, block, 0, temp.Length);
        }
    }
}
