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
using System.Collections.Generic;

namespace NPOI.POIFS.Storage
{
    /// <summary>
    /// Storage for documents that are too small to use regular
    /// DocumentBlocks for their data
    /// @author  Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class SmallDocumentBlock : BlockWritable, ListManagedBlock
    {

        private const int BLOCK_SHIFT = 6;

        private byte[]            _data;
        private const byte _default_fill         = ( byte ) 0xff;
        private const int  _block_size           =  1 << BLOCK_SHIFT;
        private const int BLOCK_MASK = _block_size - 1;
        private static int _blocks_per_big_block =
            POIFSConstants.BIG_BLOCK_SIZE / _block_size;
        private POIFSBigBlockSize _bigBlockSize;

        public SmallDocumentBlock(POIFSBigBlockSize bigBlockSize, byte[] data, int index)
        {
            _bigBlockSize = bigBlockSize;
            _blocks_per_big_block = GetBlocksPerBigBlock(bigBlockSize);
            _data = new byte[_block_size];

            System.Array.Copy(data, index*_block_size, _data, 0, _block_size);
        }

        public SmallDocumentBlock(POIFSBigBlockSize bigBlockSize)
        {
            _bigBlockSize = bigBlockSize;
            _blocks_per_big_block = GetBlocksPerBigBlock(bigBlockSize);
            _data = new byte[_block_size];
        }

        private static int GetBlocksPerBigBlock(POIFSBigBlockSize bigBlockSize)
        {
            return bigBlockSize.GetBigBlockSize()/_block_size;
        }
        /// <summary>
        /// convert a single long array into an array of SmallDocumentBlock
        /// instances
        /// </summary>
        /// <param name="bigBlockSize">the poifs bigBlockSize</param>
        /// <param name="array">the byte array to be converted</param>
        /// <param name="size">the intended size of the array (which may be smaller)</param>
        /// <returns>an array of SmallDocumentBlock instances, filled from
        /// the array</returns>
        public static SmallDocumentBlock [] Convert(POIFSBigBlockSize bigBlockSize,
                                                    byte [] array,
                                                    int size)
        {
            SmallDocumentBlock[] rval = new SmallDocumentBlock[ (size + _block_size - 1) / _block_size ];
            int                  offset = 0;

            for (int k = 0; k < rval.Length; k++)
            {
                rval[ k ] = new SmallDocumentBlock(bigBlockSize);
                if (offset < array.Length)
                {
                    int length = Math.Min(_block_size, array.Length - offset);

                    Array.Copy(array, offset, rval[ k ]._data, 0, length);
                    if (length != _block_size)
                    {
                        for (int i = length; i < _block_size; i++)
                            rval[k]._data[i] = _default_fill;
                    }
                }
                else
                {
                    for (int j = 0; j < rval[k]._data.Length; j++)
                    {
                        rval[k]._data[j] = _default_fill;
                    }

                }
                offset += _block_size;
            }
            return rval;
        }

        /// <summary>
        /// fill out a List of SmallDocumentBlocks so that it fully occupies
        /// a Set of big blocks
        /// </summary>
        /// <param name="bigBlockSize"></param>
        /// <param name="blocks">the List to be filled out.</param>
        /// <returns>number of big blocks the list encompasses</returns>
        public static int Fill(POIFSBigBlockSize bigBlockSize,IList blocks)
        {
            int _blocks_per_big_block = GetBlocksPerBigBlock(bigBlockSize);
            int count           = blocks.Count;
            int big_block_count = (count + _blocks_per_big_block - 1) / _blocks_per_big_block;
            int full_count      = big_block_count * _blocks_per_big_block;

            for (; count < full_count; count++)
            {
                blocks.Add(MakeEmptySmallDocumentBlock(bigBlockSize));
            }
            return big_block_count;
        }

        /// <summary>
        /// Factory for creating SmallDocumentBlocks from DocumentBlocks
        /// </summary>
        /// <param name="bigBlocksSize"></param>
        /// <param name="store">the original DocumentBlocks</param>
        /// <param name="size">the total document size</param>
        /// <returns>an array of new SmallDocumentBlocks instances</returns>
        public static SmallDocumentBlock [] Convert(POIFSBigBlockSize bigBlocksSize,
                                                    BlockWritable [] store,
                                                    int size)
        {
            using (MemoryStream stream = new MemoryStream())
            {

                for (int j = 0; j < store.Length; j++)
                {
                    store[j].WriteBlocks(stream);
                }
                byte[] data = stream.ToArray();
            SmallDocumentBlock[] rval = new SmallDocumentBlock[ ConvertToBlockCount(size) ];

                for (int index = 0; index < rval.Length; index++)
                {
                rval[index] = new SmallDocumentBlock(bigBlocksSize,data, index);
                }
                return rval;
            }
        }

        /// <summary>
        /// create a list of SmallDocumentBlock's from raw data
        /// </summary>
        /// <param name="bigBlockSize"></param>
        /// <param name="blocks">the raw data containing the SmallDocumentBlock</param>
        /// <returns>a List of SmallDocumentBlock's extracted from the input</returns>
        public static List<SmallDocumentBlock> Extract(POIFSBigBlockSize bigBlockSize, ListManagedBlock [] blocks)
        {
            int _blocks_per_big_block = GetBlocksPerBigBlock(bigBlockSize);
            List<SmallDocumentBlock> sdbs = new List<SmallDocumentBlock>();

            for (int j = 0; j < blocks.Length; j++)
            {
                byte[] data = blocks[ j ].Data;

                for (int k = 0; k < _blocks_per_big_block; k++)
                {
                    sdbs.Add(new SmallDocumentBlock(bigBlockSize, data, k));
                }
            }
            return sdbs;
        }

        /// <summary>
        /// Read data from an array of SmallDocumentBlocks
        /// </summary>
        /// <param name="blocks">the blocks to Read from.</param>
        /// <param name="buffer">the buffer to Write the data into.</param>
        /// <param name="offset">the offset into the array of blocks to Read from</param>
        public static void Read(BlockWritable[] blocks, byte[] buffer, int offset)
        {
            int firstBlockIndex  = offset / _block_size;
            int firstBlockOffSet = offset % _block_size;
            int lastBlockIndex   = (offset + buffer.Length - 1) / _block_size;

            if (firstBlockIndex == lastBlockIndex)
            {
                Array.Copy(
                    (( SmallDocumentBlock ) blocks[ firstBlockIndex ])._data,
                    firstBlockOffSet, buffer, 0, buffer.Length);
            }
            else
            {
                int buffer_offset = 0;

                Array.Copy(
                    (( SmallDocumentBlock ) blocks[ firstBlockIndex ])._data,
                    firstBlockOffSet, buffer, buffer_offset,
                    _block_size - firstBlockOffSet);
                buffer_offset += _block_size - firstBlockOffSet;
                for (int j = firstBlockIndex + 1; j < lastBlockIndex; j++)
                {
                    Array.Copy((( SmallDocumentBlock ) blocks[ j ])._data,
                                     0, buffer, buffer_offset, _block_size);
                    buffer_offset += _block_size;
                }
                Array.Copy(
                    (( SmallDocumentBlock ) blocks[ lastBlockIndex ])._data, 0,
                    buffer, buffer_offset, buffer.Length - buffer_offset);
            }
        }

        public static DataInputBlock GetDataInputBlock(SmallDocumentBlock[] blocks, int offset)
        {
            int firstBlockIndex = offset >> BLOCK_SHIFT;

            int firstBlockOffset = offset & BLOCK_MASK;

            return new DataInputBlock(blocks[firstBlockIndex]._data, firstBlockOffset);
        }

        /// <summary>
        /// Calculate the storage size of a Set of SmallDocumentBlocks
        /// </summary>
        /// <param name="size"> number of SmallDocumentBlocks</param>
        /// <returns>total size</returns>
        public static int CalcSize(int size)
        {
            return size * _block_size;
        }

        /// <summary>
        /// Makes the empty small document block.
        /// </summary>
        /// <returns></returns>
        private static SmallDocumentBlock MakeEmptySmallDocumentBlock(POIFSBigBlockSize bigBlockSize)
        {
            SmallDocumentBlock block = new SmallDocumentBlock(bigBlockSize);
            for (int i = 0; i < block._data.Length; i++)
            {
                block._data[i] = _default_fill;
            }
            return block;
        }

        /// <summary>
        /// Converts to block count.
        /// </summary>
        /// <param name="size">The size.</param>
        /// <returns></returns>
        private static int ConvertToBlockCount(int size)
        {
            return (size + _block_size - 1) / _block_size;
        }

        /// <summary>
        /// Write the storage to an OutputStream
        /// </summary>
        /// <param name="stream">the OutputStream to which the stored data should
        /// be written</param>
        public void WriteBlocks(Stream stream)

        {
            stream.Write(_data,0,_data.Length);
        }


        /// <summary>
        /// Get the data from the block
        /// </summary>
        /// <value>the block's data as a byte array</value>
        public byte [] Data
        {
            get
            {
                return _data;
            }
        }

        public POIFSBigBlockSize BigBlockSize
        {
            get { return _bigBlockSize; }
        }
    }
}
