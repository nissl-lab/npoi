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
using System.Collections;
using System.IO;

using NPOI.POIFS.Common;
using NPOI.Util;

namespace NPOI.POIFS.Storage
{
    /// <summary>
    /// The block containing the archive header
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class HeaderBlockWriter : BigBlock
    {
        private static byte _default_value = (byte)0xFF;

        // number of big block allocation table blocks (int)
        private IntegerField      _bat_count;

        // start of the property  block (int index of the property 
        // chain's first big block)
        private IntegerField      _property_start;

        // start of the small block allocation table (int index of small
        // block allocation table's first big block)
        private IntegerField      _sbat_start;

        // number of big blocks holding the small block allocation table
        private IntegerField      _sbat_block_count;

        // big block index for extension to the big block allocation table
        private IntegerField      _xbat_start;
        private IntegerField      _xbat_count;
        private byte[]            _data;

        /// <summary>
        /// Create a single instance initialized with default values
        /// </summary>
        public HeaderBlockWriter()
        {
            _data = new byte[ POIFSConstants.BIG_BLOCK_SIZE ];
            for (int i = 0; i < this._data.Length; i++)
            {
                this._data[i] = _default_value;
            }
            new LongField(HeaderBlockConstants._signature_offset, HeaderBlockConstants._signature, _data);
            new IntegerField(0x08, 0, _data);
            new IntegerField(0x0c, 0, _data);
            new IntegerField(0x10, 0, _data);
            new IntegerField(0x14, 0, _data);
            new ShortField(0x18, ( short ) 0x3b, ref _data);
            new ShortField(0x1a, ( short ) 0x3, ref _data);
            new ShortField(0x1c, ( short ) -2, ref _data);
            new ShortField(0x1e, ( short ) 0x9, ref _data);
            new IntegerField(0x20, 0x6, _data);
            new IntegerField(0x24, 0, _data);
            new IntegerField(0x28, 0, _data);
            _bat_count = new IntegerField(HeaderBlockConstants._bat_count_offset, 0, _data);
            _property_start = new IntegerField(HeaderBlockConstants._property_start_offset,
                                               POIFSConstants.END_OF_CHAIN,
                                               _data);
            new IntegerField(0x34, 0, _data);
            new IntegerField(0x38, 0x1000, _data);
            _sbat_start = new IntegerField(HeaderBlockConstants._sbat_start_offset,
                                           POIFSConstants.END_OF_CHAIN, _data);
            _sbat_block_count = new IntegerField(HeaderBlockConstants._sbat_block_count_offset, 0,
					         _data);
            _xbat_start = new IntegerField(HeaderBlockConstants._xbat_start_offset,
                                           POIFSConstants.END_OF_CHAIN, _data);
            _xbat_count = new IntegerField(HeaderBlockConstants._xbat_count_offset, 0, _data);
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
        public BATBlock [] SetBATBlocks(int blockCount,
                                        int startBlock)
        {
            BATBlock[] rvalue;

            _bat_count.Set(blockCount, _data);
            int limit = Math.Min(blockCount, HeaderBlockConstants._max_bats_in_header);
            int offset = HeaderBlockConstants._bat_array_offset;

            for (int j = 0; j < limit; j++)
            {
                new IntegerField(offset, startBlock + j, _data);
                offset += LittleEndianConstants.INT_SIZE;
            }
            if (blockCount > HeaderBlockConstants._max_bats_in_header)
            {
                int excess_blocks = blockCount - HeaderBlockConstants._max_bats_in_header;
                int[] excess_block_array = new int[ excess_blocks ];

                for (int j = 0; j < excess_blocks; j++)
                {
                    excess_block_array[ j ] = startBlock + j
                                              + HeaderBlockConstants._max_bats_in_header;
                }
                rvalue = BATBlock.CreateXBATBlocks(excess_block_array,
                                                   startBlock + blockCount);
                _xbat_start.Set(startBlock + blockCount, _data);
            }
            else
            {
                rvalue = BATBlock.CreateXBATBlocks(new int[ 0 ], 0);
                _xbat_start.Set(POIFSConstants.END_OF_CHAIN, _data);
            }
            _xbat_count.Set(rvalue.Length, _data);
            return rvalue;
        }

        /// <summary>
        /// Set start of Property Table
        /// </summary>
        /// <value>the index of the first block of the Property
        /// Table</value>
        public int PropertyStart
        {
            set { _property_start.Set(value, _data); }
        }

        /// <summary>
        /// Set start of small block allocation table
        /// </summary>
        /// <value>the index of the first big block of the small
        /// block allocation table</value>
        public int SBATStart
        {
            set { _sbat_start.Set(value, _data); }

        }

        /// <summary>
        /// Set count of SBAT blocks
        /// </summary>
        /// <value>the number of SBAT blocks</value>
        public int SBATBlockCount
        {
            get{return _sbat_block_count.Value;}
	        set{_sbat_block_count.Set(value, _data);}
        }

        /// <summary>
        /// For a given number of BAT blocks, calculate how many XBAT
        /// blocks will be needed
        /// </summary>
        /// <param name="blockCount">number of BAT blocks</param>
        /// <returns>number of XBAT blocks needed</returns>
        public static int CalculateXBATStorageRequirements(int blockCount)
        {
            return (blockCount > HeaderBlockConstants._max_bats_in_header)
                   ? BATBlock.CalculateXBATStorageRequirements(blockCount
                       - HeaderBlockConstants._max_bats_in_header)
                   : 0;
        }


        /// <summary>
        /// Write the block's data to an Stream
        /// </summary>
        /// <param name="stream">the Stream to which the stored data should
        /// be written
        /// </param>
        internal override void WriteData(Stream stream)
        {
            WriteData(stream, _data);
        }
    }
}
