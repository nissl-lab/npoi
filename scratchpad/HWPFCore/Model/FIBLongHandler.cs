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
namespace NPOI.HWPF.Model
{

    /**
     * Handles the fibRgLw / The FibRgLw97 part of
     *  the FIB (File Information Block)
     */
    public class FIBLongHandler
    {
        public const int CBMAC = 0;
        public const int PRODUCTCREATED = 1;
        public const int PRODUCTREVISED = 2;
        public const int CCPTEXT = 3;
        public const int CCPFTN = 4;
        public const int CCPHDD = 5;
        public const int CCPMCR = 6;
        public const int CCPATN = 7;
        public const int CCPEDN = 8;
        public const int CCPTXBX = 9;
        public const int CCPHDRTXBX = 10;
        public const int PNFBPCHPFIRST = 11;
        public const int PNCHPFIRST = 12;
        public const int CPNBTECHP = 13;
        public const int PNFBPPAPFIRST = 14;
        public const int PNPAPFIRST = 15;
        public const int CPNBTEPAP = 16;
        public const int PNFBPLVCFIRST = 17;
        public const int PNLVCFIRST = 18;
        public const int CPNBTELVC = 19;
        public const int FCISLANDFIRST = 20;
        public const int FCISLANDLIM = 21;

        int[] _longs;

        public FIBLongHandler(byte[] mainStream, int offset)
        {
            int longCount = LittleEndian.GetShort(mainStream, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _longs = new int[longCount];

            for (int x = 0; x < longCount; x++)
            {
                _longs[x] = LittleEndian.GetInt(mainStream, offset + (x * LittleEndianConsts.INT_SIZE));
            }
        }

        /**
         * Refers to a 32 bit windows "long" same as a Java int
         * @param longCode FIXME: Document this!
         * @return FIXME: Document this!
         */
        public int GetLong(int longCode)
        {
            return _longs[longCode];
        }

        public void SetLong(int longCode, int value)
        {
            _longs[longCode] = value;
        }

        internal void Serialize(byte[] mainStream, int offset)
        {
            LittleEndian.PutShort(mainStream, offset, (short)_longs.Length);
            offset += LittleEndianConsts.SHORT_SIZE;

            for (int x = 0; x < _longs.Length; x++)
            {
                LittleEndian.PutInt(mainStream, offset, _longs[x]);
                offset += LittleEndianConsts.INT_SIZE;
            }
        }

        internal int SizeInBytes()
        {
            return (_longs.Length * LittleEndianConsts.INT_SIZE) + LittleEndianConsts.SHORT_SIZE;
        }


    }
}

