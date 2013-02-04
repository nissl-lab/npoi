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
    public class DocumentBlock : BigBlock
    {
        private static byte _default_value = (byte)0xFF;
        private byte[] _data;
        private int _bytes_Read;

        /// <summary>
        /// create a document block from a raw data block
        /// </summary>
        /// <param name="block">The block.</param>
        public DocumentBlock(RawDataBlock block) : 
			base(block.BigBlockSize == POIFSConstants.SMALLER_BIG_BLOCK_SIZE ? 
                    POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS : POIFSConstants.LARGER_BIG_BLOCK_SIZE_DETAILS)
        {
            _data = block.Data;
            _bytes_Read = _data.Length;
        }

        /// <summary>
        /// Create a single instance initialized with data.
        /// </summary>
        /// <param name="stream">the InputStream delivering the data.</param>
        /// <param name="bigBlockSize">the poifs bigBlockSize</param>
        public DocumentBlock(Stream stream, POIFSBigBlockSize bigBlockSize)
            : this(bigBlockSize)
        {
            int count = IOUtils.ReadFully(stream, _data);
            _bytes_Read = (count == -1) ? 0: count;
        }

        public DocumentBlock(POIFSBigBlockSize bigBlockSize)
            :base(bigBlockSize)
        {
            _data = new byte[POIFSConstants.BIG_BLOCK_SIZE];
            Arrays.Fill(_data, _default_value);
        }

        /// <summary>
        /// Get the number of bytes Read for this block.
        /// </summary>
        /// <value>bytes Read into the block</value>
        public int Size
        {
            get { return _bytes_Read; }
        }

        /// <summary>
        /// Was this a partially Read block?
        /// </summary>
        /// <value><c>true</c> if the block was only partially filled with data</value>
        public bool PartiallyRead
        {
            get { return _bytes_Read != POIFSConstants.BIG_BLOCK_SIZE; }
        }

        /// <summary>
        /// Gets the fill byte used
        /// </summary>
        /// <value>The fill byte.</value>
        public static byte FillByte
        {
            get { return _default_value; }
        }

        /// <summary>
        /// convert a single long array into an array of DocumentBlock
        /// instances
        /// </summary>
        /// <param name="bigBlockSize">the poifs bigBlockSize</param>
        /// <param name="array">the byte array to be converted</param>
        /// <param name="size">the intended size of the array (which may be smaller)</param>
        /// <returns>an array of DocumentBlock instances, filled from the
        /// input array</returns>
        public static DocumentBlock[] Convert(POIFSBigBlockSize bigBlockSize,
                                                byte[] array,
                                               int size)
        {
            DocumentBlock[] rval =
                new DocumentBlock[(size + POIFSConstants.BIG_BLOCK_SIZE - 1) / POIFSConstants.BIG_BLOCK_SIZE];
            int offset = 0;

            for (int k = 0; k < rval.Length; k++)
            {
                rval[k] = new DocumentBlock(bigBlockSize);
                if (offset < array.Length)
                {
                    int length = Math.Min(POIFSConstants.BIG_BLOCK_SIZE,
                                          array.Length - offset);

                    Array.Copy(array, offset, rval[k]._data, 0, length);
                    if (length != POIFSConstants.BIG_BLOCK_SIZE)
                    {
                        for (int j = (length > 0) ? (length - 1) : length; j < POIFSConstants.BIG_BLOCK_SIZE; j++)
                        {
                            rval[k]._data[j] = _default_value;
                        }
                    }
                }
                else
                {
                    for (int j = 0; j < rval[k]._data.Length; j++)
                    {
                        rval[k]._data[j] = _default_value;
                    }
                }
                offset += POIFSConstants.BIG_BLOCK_SIZE;
            }
            return rval;
        }

        /// <summary>
        /// Read data from an array of DocumentBlocks
        /// </summary>
        /// <param name="blocks">the blocks to Read from</param>
        /// <param name="buffer">the buffer to Write the data into</param>
        /// <param name="offset">the offset into the array of blocks to Read from</param>
        public static void Read(DocumentBlock[] blocks,
                                byte[] buffer, int offset)
        {
            int firstBlockIndex = offset / POIFSConstants.BIG_BLOCK_SIZE;
            int firstBlockOffSet = offset % POIFSConstants.BIG_BLOCK_SIZE;
            int lastBlockIndex = (offset + buffer.Length - 1)
                                   / POIFSConstants.BIG_BLOCK_SIZE;

            if (firstBlockIndex == lastBlockIndex)
            {
                Array.Copy(blocks[firstBlockIndex]._data,
                                 firstBlockOffSet, buffer, 0, buffer.Length);
            }
            else
            {
                int buffer_offset = 0;

                Array.Copy(blocks[firstBlockIndex]._data,
                                 firstBlockOffSet, buffer, buffer_offset,
                                 POIFSConstants.BIG_BLOCK_SIZE
                                 - firstBlockOffSet);
                buffer_offset += POIFSConstants.BIG_BLOCK_SIZE - firstBlockOffSet;
                for (int j = firstBlockIndex + 1; j < lastBlockIndex; j++)
                {
                    Array.Copy(blocks[j]._data, 0, buffer, buffer_offset,
                                     POIFSConstants.BIG_BLOCK_SIZE);
                    buffer_offset += POIFSConstants.BIG_BLOCK_SIZE;
                }
                Array.Copy(blocks[lastBlockIndex]._data, 0, buffer,
                                 buffer_offset, buffer.Length - buffer_offset);
            }
        }

        public static DataInputBlock GetDataInputBlock(DocumentBlock[] blocks, int offset)
        {
            if (blocks == null || blocks.Length == 0)
                return null;

            POIFSBigBlockSize bigBlockSize = blocks[0].bigBlockSize;
            int BLOCK_SHIFT = bigBlockSize.GetHeaderValue();
            int BLOCK_SIZE = bigBlockSize.GetBigBlockSize();
            int BLOCK_MASK = BLOCK_SIZE - 1;

            int firstBlockIndex = offset >> BLOCK_SHIFT;
            int firstBlockOffset = offset & BLOCK_MASK;
            return new DataInputBlock(blocks[firstBlockIndex]._data, firstBlockOffset);
        }
        /// <summary>
        /// Write the storage to an OutputStream
        /// </summary>
        /// <param name="stream">the OutputStream to which the stored data should
        /// be written</param>
        public override void WriteData(Stream stream)
        {
            WriteData(stream, _data);
        }

    }
}