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
    /// A block of block allocation table entries. BATBlocks are created
    /// only through a static factory method: createBATBlocks.
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class BATBlock:BigBlock
    {
        private static int  _entries_per_block      =
            POIFSConstants.BIG_BLOCK_SIZE / LittleEndianConstants.INT_SIZE;
        private static int  _entries_per_xbat_block = _entries_per_block - 1;
        private static int  _xbat_chain_offset      =
            _entries_per_xbat_block * LittleEndianConstants.INT_SIZE;
        private static byte _default_value = ( byte ) 0xFF;
        private IntegerField[]    _fields;
        private byte[]            _data;

        public override void Dispose()
        {
            _data = null;
            _fields = null;

        }
        /// <summary>
        /// Create a single instance initialized with default values
        /// </summary>
        private BATBlock()
        {
            _data = new byte[ POIFSConstants.BIG_BLOCK_SIZE ];
            for (int i = 0; i < this._data.Length; i++)
            {
                this._data[i] = _default_value;
            }
            _fields = new IntegerField[ _entries_per_block ];
            int offset = 0;

            for (int j = 0; j < _entries_per_block; j++)
            {
                _fields[ j ] = new IntegerField(offset);
                offset       += LittleEndianConstants.INT_SIZE;
            }
        }

        /// <summary>
        /// Create an array of BATBlocks from an array of int block
        /// allocation table entries
        /// </summary>
        /// <param name="entries">the array of int entries</param>
        /// <returns>the newly created array of BATBlocks</returns>
        public static BATBlock [] CreateBATBlocks(int [] entries)
        {
            int        block_count = CalculateStorageRequirements(entries.Length);
            BATBlock[] blocks      = new BATBlock[ block_count ];
            int        index       = 0;
            int        remaining   = entries.Length;

            for (int j = 0; j < entries.Length; j += _entries_per_block)
            {
                blocks[ index++ ] = new BATBlock(entries, j,
                                                 (remaining > _entries_per_block)
                                                 ? j + _entries_per_block
                                                 : entries.Length);
                remaining         -= _entries_per_block;
            }
            return blocks;
        }

        /// <summary>
        /// Create an array of XBATBlocks from an array of int block
        /// allocation table entries
        /// </summary>
        /// <param name="entries">the array of int entries</param>
        /// <param name="startBlock">the start block of the array of XBAT blocks</param>
        /// <returns>the newly created array of BATBlocks</returns>
        public static BATBlock [] CreateXBATBlocks(int [] entries,
                                                   int startBlock)
        {
            int        block_count =
                CalculateXBATStorageRequirements(entries.Length);
            BATBlock[] blocks      = new BATBlock[ block_count ];
            int        index       = 0;
            int        remaining   = entries.Length;

            if (block_count != 0)
            {
                for (int j = 0; j < entries.Length; j += _entries_per_xbat_block)
                {
                    blocks[ index++ ] =
                        new BATBlock(entries, j,
                                     (remaining > _entries_per_xbat_block)
                                     ? j + _entries_per_xbat_block
                                     : entries.Length);
                    remaining         -= _entries_per_xbat_block;
                }
                for (index = 0; index < blocks.Length - 1; index++)
                {
                    blocks[ index ].XBATChain=startBlock + index + 1;
                }
                blocks[ index ].XBATChain=POIFSConstants.END_OF_CHAIN;
            }
            return blocks;
        }

        /// <summary>
        /// Calculate how many BATBlocks are needed to hold a specified
        /// number of BAT entries.
        /// </summary>
        /// <param name="entryCount">the number of entries</param>
        /// <returns>the number of BATBlocks needed</returns>
        public static int CalculateStorageRequirements(int entryCount)
        {
            return (entryCount + _entries_per_block - 1) / _entries_per_block;
        }

        /// <summary>
        /// Calculate how many XBATBlocks are needed to hold a specified
        /// number of BAT entries.
        /// </summary>
        /// <param name="entryCount">the number of entries</param>
        /// <returns>the number of XBATBlocks needed</returns>
        public static int CalculateXBATStorageRequirements(int entryCount)
        {
            return (entryCount + _entries_per_xbat_block - 1)
                   / _entries_per_xbat_block;
        }

        /// <summary>
        /// Gets the entries per block.
        /// </summary>
        /// <value>The number of entries per block</value>
        public static int EntriesPerBlock
        {
            get{return _entries_per_block;}
        }

        /// <summary>
        /// Gets the entries per XBAT block.
        /// </summary>
        /// <value>number of entries per XBAT block</value>
        public static int EntriesPerXBATBlock
        {
            get{return _entries_per_xbat_block;}
        }

        /// <summary>
        /// Gets the XBAT chain offset.
        /// </summary>
        /// <value>offset of chain index of XBAT block</value>
        public static int XBATChainOffset
        {
            get{return _xbat_chain_offset;}
        }

        private int XBATChain
        {
            set { _fields[_entries_per_xbat_block].Set(value, _data); }
        }

        /// <summary>
        /// Create a single instance initialized (perhaps partially) with entries
        /// </summary>
        /// <param name="entries">the array of block allocation table entries</param>
        /// <param name="start_index">the index of the first entry to be written
        /// to the block</param>
        /// <param name="end_index">the index, plus one, of the last entry to be
        /// written to the block (writing is for all index
        /// k, start_index less than k less than end_index)
        /// </param>
        private BATBlock(int [] entries, int start_index,
                         int end_index):this()
        {
            
            for (int k = start_index; k < end_index; k++)
            {
                _fields[ k - start_index ].Set(entries[ k ], _data);
            }
        }

        /// <summary>
        /// Write the block's data to an Stream
        /// </summary>
        /// <param name="stream">the Stream to which the stored data should
        /// be written</param>
        internal override void WriteData(Stream stream)
        {
            WriteData(stream, _data);
        }
    }
}
