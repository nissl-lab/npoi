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
using System.Collections.Generic;

using NPOI.POIFS.Common;
using NPOI.Util;

namespace NPOI.POIFS.Storage
{
    /// <summary>
    /// A block of block allocation table entries. BATBlocks are created
    /// only through a static factory method: createBATBlocks.
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class BATBlock : BigBlock
    {
        private static int _entries_per_block =
            POIFSConstants.BIG_BLOCK_SIZE / LittleEndianConsts.INT_SIZE;
        private static int _entries_per_xbat_block = _entries_per_block - 1;
        private static int _xbat_chain_offset =
            _entries_per_xbat_block * LittleEndianConsts.INT_SIZE;
        private static byte _default_value = (byte)0xFF;
        private IntegerField[] _fields;
        private byte[] _data;
        /**
         * For a regular fat block, these are 128 / 1024 
         *  next sector values.
         * For a XFat (DIFat) block, these are 127 / 1023
         *  next sector values, then a chaining value.
         */
        private int[] _values;

        /**
         * Does this BATBlock have any free sectors in it?
         */
        private bool _has_free_sectors;

        /**
         * Where in the file are we?
         */
        private int ourBlockIndex;

        /// <summary>
        /// Create a single instance initialized with default values
        /// </summary>
        protected BATBlock()
        {
            _data = new byte[POIFSConstants.BIG_BLOCK_SIZE];
            for (int i = 0; i < this._data.Length; i++)
            {
                this._data[i] = _default_value;
            }
            _fields = new IntegerField[_entries_per_block];
            int offset = 0;

            for (int j = 0; j < _entries_per_block; j++)
            {
                _fields[j] = new IntegerField(offset);
                offset += LittleEndianConsts.INT_SIZE;
            }
        }
        protected BATBlock(POIFSBigBlockSize bigBlockSize)
            : base(bigBlockSize)
        {
            int _entries_per_block = bigBlockSize.GetBATEntriesPerBlock();
            _values = new int[_entries_per_block];
            _has_free_sectors = true;
         
            for (int i = 0; i < _values.Length; i++)
                _values[i] = POIFSConstants.UNUSED_BLOCK;
        }
        /**
         * Create a single instance initialized (perhaps partially) with entries
         *
         * @param entries the array of block allocation table entries
         * @param start_index the index of the first entry to be written
         *                    to the block
         * @param end_index the index, plus one, of the last entry to be
         *                  written to the block (writing is for all index
         *                  k, start_index &lt;= k &lt; end_index)
         */

        protected BATBlock(POIFSBigBlockSize bigBlockSize, int[] entries,
                         int start_index, int end_index)
            : this(bigBlockSize)
        {

            for (int k = start_index; k < end_index; k++)
            {
                _values[k - start_index] = entries[k];
            }

            // Do we have any free sectors?
            if (end_index - start_index == _values.Length)
            {
                RecomputeFree();
            }
        }
        private void RecomputeFree()
        {
            bool hasFree = false;
            for (int k = 0; k < _values.Length; k++)
            {
                if (_values[k] == POIFSConstants.UNUSED_BLOCK)
                {
                    hasFree = true;
                    break;
                }
            }
            _has_free_sectors = hasFree;
        }
        /**
         * Create a single BATBlock from the byte buffer, which must hold at least
         *  one big block of data to be read.
         */
        public static BATBlock CreateBATBlock(POIFSBigBlockSize bigBlockSize, BinaryReader data)
        {
            // Create an empty block
            BATBlock block = new BATBlock(bigBlockSize);

            // Fill it
            byte[] buffer = new byte[LittleEndianConsts.INT_SIZE];
            for (int i = 0; i < block._values.Length; i++)
            {
                data.Read(buffer,0,buffer.Length);
                block._values[i] = LittleEndian.GetInt(buffer);
            }
            block.RecomputeFree();

            // All done
            return block;
        }

        //public static BATBlock CreateBATBlock(POIFSBigBlockSize bigBlockSize, byte[] data)
        public static BATBlock CreateBATBlock(POIFSBigBlockSize bigBlockSize, ByteBuffer data)
        {
            // Create an empty block
            BATBlock block = new BATBlock(bigBlockSize);

            // Fill it
            byte[] buffer = new byte[LittleEndianConsts.INT_SIZE];
            //int index = 0;
            for (int i = 0; i < block._values.Length; i++)
            {
                //data.Read(buffer, 0, buffer.Length);
                //for (int j = 0; j < buffer.Length; j++, index++)
                 //   buffer[j] = data[index];
                data.Read(buffer);
                block._values[i] = LittleEndian.GetInt(buffer);
            }
            block.RecomputeFree();

            // All done
            return block;
        }
        ///**
        // * Creates a single BATBlock, with all the values set to empty.
        // */
        public static BATBlock CreateEmptyBATBlock(POIFSBigBlockSize bigBlockSize, bool isXBAT)
        {
            BATBlock block = new BATBlock(bigBlockSize);
            if (isXBAT)
            {
                block.SetXBATChain(bigBlockSize, POIFSConstants.END_OF_CHAIN);
            }
            return block;
        }


        /// <summary>
        /// Create an array of BATBlocks from an array of int block
        /// allocation table entries
        /// </summary>
        /// <param name="bigBlockSize">the poifs bigBlockSize</param>
        /// <param name="entries">the array of int entries</param>
        /// <returns>the newly created array of BATBlocks</returns>
        public static BATBlock[] CreateBATBlocks(POIFSBigBlockSize bigBlockSize, int[] entries)
        {
            int block_count = CalculateStorageRequirements(entries.Length);
            BATBlock[] blocks = new BATBlock[block_count];
            int index = 0;
            int remaining = entries.Length;

            for (int j = 0; j < entries.Length; j += _entries_per_block)
            {
                blocks[index++] = new BATBlock(bigBlockSize, entries, j,
                                                 (remaining > _entries_per_block)
                                                 ? j + _entries_per_block
                                                 : entries.Length);
                remaining -= _entries_per_block;
            }
            return blocks;
        }

        /// <summary>
        /// Create an array of XBATBlocks from an array of int block
        /// allocation table entries
        /// </summary>
        /// <param name="bigBlockSize"></param>
        /// <param name="entries">the array of int entries</param>
        /// <param name="startBlock">the start block of the array of XBAT blocks</param>
        /// <returns>the newly created array of BATBlocks</returns>
        public static BATBlock[] CreateXBATBlocks(POIFSBigBlockSize bigBlockSize, int[] entries,
                                                   int startBlock)
        {
            int block_count =
                CalculateXBATStorageRequirements(entries.Length);
            BATBlock[] blocks = new BATBlock[block_count];
            int index = 0;
            int remaining = entries.Length;

            if (block_count != 0)
            {
                for (int j = 0; j < entries.Length; j += _entries_per_xbat_block)
                {
                    blocks[index++] =
                        new BATBlock(bigBlockSize, entries, j,
                                     (remaining > _entries_per_xbat_block)
                                     ? j + _entries_per_xbat_block
                                     : entries.Length);
                    remaining -= _entries_per_xbat_block;
                }
                for (index = 0; index < blocks.Length - 1; index++)
                {
                    blocks[index].SetXBATChain(bigBlockSize, startBlock + index + 1);
                }
                blocks[index].SetXBATChain(bigBlockSize, POIFSConstants.END_OF_CHAIN);
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

        public static int CalculateStorageRequirements(POIFSBigBlockSize bigBlockSize, int entryCount)
        {
            int _entries_per_block = bigBlockSize.GetBATEntriesPerBlock();
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

        public static int CalculateXBATStorageRequirements(POIFSBigBlockSize bigBlockSize, int entryCount)
        {
            int _entries_per_xbat_block = bigBlockSize.GetXBATEntriesPerBlock();

            return (entryCount + _entries_per_xbat_block - 1) / _entries_per_xbat_block;
        }
        /**
         * Calculates the maximum size of a file which is addressable given the
         *  number of FAT (BAT) sectors specified. (We don't care if those BAT
         *  blocks come from the 109 in the header, or from header + XBATS, it
         *  won't affect the calculation)
         *  
         * The actual file size will be between [size of fatCount-1 blocks] and
         *   [size of fatCount blocks].
         *  For 512 byte block sizes, this means we may over-estimate by up to 65kb.
         *  For 4096 byte block sizes, this means we may over-estimate by up to 4mb
         */
        public static long CalculateMaximumSize(POIFSBigBlockSize bigBlockSize,
              int numBATs)
        {
            long size = 1; // Header isn't FAT addressed

            // The header has up to 109 BATs, and extra ones are referenced
            //  from XBATs
            // However, all BATs can contain 128/1024 blocks
            size += (numBATs * bigBlockSize.GetBATEntriesPerBlock());

            // So far we've been in sector counts, turn into bytes
            return size * bigBlockSize.GetBigBlockSize();
        }
        public static long CalculateMaximumSize(HeaderBlock header)
        {
            return CalculateMaximumSize(header.BigBlockSize, header.BATCount);
        }

        public static BATBlockAndIndex GetBATBlockAndIndex(int offset, HeaderBlock header, List<BATBlock> bats)
        {
            POIFSBigBlockSize bigBlockSize = header.BigBlockSize;

            int whichBAT = (int)Math.Floor(1.0*offset / bigBlockSize.GetBATEntriesPerBlock());
            int index = offset % bigBlockSize.GetBATEntriesPerBlock();
            return new BATBlockAndIndex(index, bats[whichBAT]);
        }

        public static BATBlockAndIndex GetSBATBlockAndIndex(int offset, HeaderBlock header, List<BATBlock> sbats)
        {
            POIFSBigBlockSize bigBlockSize = header.BigBlockSize;

            int whichSBAT = (int)Math.Floor(1.0*offset / bigBlockSize.GetBATEntriesPerBlock());
            int index = offset % bigBlockSize.GetBATEntriesPerBlock();

            return new BATBlockAndIndex(index, sbats[whichSBAT]);
        }

        /// <summary>
        /// Gets the entries per block.
        /// </summary>
        /// <value>The number of entries per block</value>
        public static int EntriesPerBlock
        {
            get { return _entries_per_block; }
        }

        /// <summary>
        /// Gets the entries per XBAT block.
        /// </summary>
        /// <value>number of entries per XBAT block</value>
        public static int EntriesPerXBATBlock
        {
            get { return _entries_per_xbat_block; }
        }

        /// <summary>
        /// Gets the XBAT chain offset.
        /// </summary>
        /// <value>offset of chain index of XBAT block</value>
        public static int XBATChainOffset
        {
            get { return _xbat_chain_offset; }
        }

        private void SetXBATChain(int chainIndex)
        {
            _fields[_entries_per_xbat_block].Set(chainIndex, _data);
        }

        private void SetXBATChain(POIFSBigBlockSize bigBlockSize, int chainIndex)
        {
            int _entries_per_xbat_block = bigBlockSize.GetXBATEntriesPerBlock();
            _values[_entries_per_xbat_block] = chainIndex;
        }
        /**
         * Does this BATBlock have any free sectors in it, or
         *  is it full?
         */
        public bool HasFreeSectors
        {
            get
            {
                return _has_free_sectors;
            }
        }

        public int GetValueAt(int relativeOffset)
        {
            if (relativeOffset >= _values.Length)
            {
                throw new IndexOutOfRangeException(
                      "Unable to fetch offset " + relativeOffset + " as the " +
                      "BAT only contains " + _values.Length + " entries"
                );
            }
            return _values[relativeOffset];
        }
        public void SetValueAt(int relativeOffset, int value)
        {
            int oldValue = _values[relativeOffset];
            _values[relativeOffset] = value;

            // Do we need to re-compute the free?
            if (value == POIFSConstants.UNUSED_BLOCK)
            {
                _has_free_sectors = true;
                return;
            }
            if (oldValue == POIFSConstants.UNUSED_BLOCK)
            {
                RecomputeFree();
            }
        }

        /**
         * Retrieve where in the file we live 
         */
        public int OurBlockIndex
        {
            get
            {
                return ourBlockIndex;
            }
            set
            {
                this.ourBlockIndex = value;
            }
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
        private BATBlock(int[] entries, int start_index,
                         int end_index)
            : this()
        {

            for (int k = start_index; k < end_index; k++)
            {
                _fields[k - start_index].Set(entries[k], _data);
            }
        }

        public void WriteData(ByteBuffer block)
        {
            block.Write(Serialize());
        }

        /// <summary>
        /// Write the block's data to an Stream
        /// </summary>
        /// <param name="stream">the Stream to which the stored data should
        /// be written</param>
        public override void WriteData(Stream stream)
        {
            byte[] buff = Serialize();
            stream.Write(buff, 0, buff.Length);
        }

        public void WriteData(byte[] block)
        {
            byte[] data = Serialize();
            for (int i = 0; i < data.Length; i++)
                block[i] = data[i];
        }

        private byte[] Serialize()
        {
            byte[] data = new byte[bigBlockSize.GetBigBlockSize()];

            int offset = 0;
            for (int i = 0; i < _values.Length; i++)
            {
                LittleEndian.PutInt(data, offset, _values[i]);
                offset += LittleEndianConsts.INT_SIZE;
            }

            return data;
        }
    }
    public class BATBlockAndIndex
    {
       private int index;
       private BATBlock block;

        public BATBlockAndIndex(int index, BATBlock block)
        {
          this.index = index;
          this.block = block;
       }
       public int Index 
       {
           get { return index; }
       }
       public BATBlock Block
       {
           get { return block; }
       }
    }
}