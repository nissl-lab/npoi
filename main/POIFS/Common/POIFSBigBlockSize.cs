
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


using NPOI.Util;
namespace NPOI.POIFS.Common
{

    /**
     * <p>A class describing attributes of the Big Block Size</p>
     */
    public class POIFSBigBlockSize
    {
        private int bigBlockSize;
        private short headerValue;

        internal POIFSBigBlockSize(int bigBlockSize, short headerValue)
        {
            this.bigBlockSize = bigBlockSize;
            this.headerValue = headerValue;
        }

        public int GetBigBlockSize()
        {
            return bigBlockSize;
        }

        /**
         * Returns the value that Gets written into the 
         *  header.
         * Is the power of two that corresponds to the
         *  size of the block, eg 512 => 9
         */
        public short GetHeaderValue()
        {
            return headerValue;
        }

        public int GetPropertiesPerBlock()
        {
            return bigBlockSize / POIFSConstants.PROPERTY_SIZE;
        }

        public int GetBATEntriesPerBlock()
        {
            return bigBlockSize / LittleEndianConsts.INT_SIZE;
        }
        public int GetXBATEntriesPerBlock()
        {
            return GetBATEntriesPerBlock() - 1;
        }
        public int GetNextXBATChainOffset()
        {
            return GetXBATEntriesPerBlock() * LittleEndianConsts.INT_SIZE;
        }
    }

}