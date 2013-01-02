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
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.IO;

using NUnit.Framework;

using NPOI.POIFS.Storage;
using NPOI.Util;
using NPOI.POIFS.Common;

namespace TestCases.POIFS.Storage
{
    /**
     * Class LocalRawDataBlockList
     *
     * @author Marc Johnson(mjohnson at apache dot org)
     */

    public class LocalRawDataBlockList : RawDataBlockList
    {
        private List<RawDataBlock> _list;
        private RawDataBlock[] _array;

        /**
         * Constructor LocalRawDataBlockList
         *
         * @exception IOException
         */

        public LocalRawDataBlockList()
            : base(new MemoryStream(new byte[0]), POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS)
        {
            _list = new List<RawDataBlock>();
            _array = null;
        }

        /**
         * create and a new XBAT block
         *
         * @param start index of first BAT block
         * @param end index of last BAT block
         * @param chain index of next XBAT block
         *
         * @exception IOException
         */

        public void CreateNewXBATBlock(int start, int end,
                                       int chain)
        {
            byte[] data = new byte[512];
            int offset = 0;

            for (int k = start; k <= end; k++)
            {
                LittleEndian.PutInt(data, offset, k);
                offset += LittleEndianConsts.INT_SIZE;
            }
            while (offset != 508)
            {
                LittleEndian.PutInt(data, offset, -1);
                offset += LittleEndianConsts.INT_SIZE;
            }
            LittleEndian.PutInt(data, offset, chain);
            Add(new RawDataBlock(new MemoryStream(data)));
        }

        /**
         * create a BAT block and Add it to the ArrayList
         *
         * @param start_index initial index for the block ArrayList
         *
         * @exception IOException
         */

        public void CreateNewBATBlock(int start_index)
        {
            byte[] data = new byte[512];
            int offset = 0;

            for (int j = 0; j < 128; j++)
            {
                int index = start_index + j;

                if (index % 256 == 0)
                {
                    LittleEndian.PutInt(data, offset, -1);
                }
                else if (index % 256 == 255)
                {
                    LittleEndian.PutInt(data, offset, -2);
                }
                else
                {
                    LittleEndian.PutInt(data, offset, index + 1);
                }
                offset += LittleEndianConsts.INT_SIZE;
            }
            Add(new RawDataBlock(new MemoryStream(data)));
        }

        /**
         * fill the ArrayList with dummy blocks
         *
         * @param count of blocks
         *
         * @exception IOException
         */

        public void Fill(int count)
        {
            int limit = 128 * count;

            for (int j = _list.Count; j < limit; j++)
            {
                Add(new RawDataBlock(new MemoryStream(new byte[0])));
            }
        }

        /**
         * Add a new block
         *
         * @param block new block to Add
         */

        public void Add(RawDataBlock block)
        {
            _list.Add(block);
        }

        /**
         * override of Remove method
         *
         * @param index of block to be Removed
         *
         * @return desired block
         *
         * @exception IOException
         */

        public override ListManagedBlock Remove(int index)
        {
            EnsureArrayExists();
            RawDataBlock rvalue = null;

            try
            {
                rvalue = _array[index];
                if (rvalue == null)
                {
                    throw new IOException("index " + index + " Is null");
                }
                _array[index] = null;
            }
            catch (IndexOutOfRangeException )
            {
                throw new IOException("Cannot Remove block[ " + index
                                      + " ]; out of range");
            }
            return (ListManagedBlock)rvalue;
        }

        /**
         * Remove the specified block from the ArrayList
         *
         * @param index the index of the specified block; if the index Is
         *              out of range, that's ok
         */

        public override void Zap(int index)
        {
            EnsureArrayExists();
            if ((index >= 0) && (index < _array.Length))
            {
                _array[index] = null;
            }
        }

        private void EnsureArrayExists()
        {
            if (_array == null)
            {
                _array = _list.ToArray();
            }
        }

        public override int BlockCount()
        {
            return _list.Count;
        }
    }
}