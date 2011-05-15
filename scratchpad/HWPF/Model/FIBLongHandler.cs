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
        public static int CBMAC = 0;
        public static int PRODUCTCREATED = 1;
        public static int PRODUCTREVISED = 2;
        public static int CCPTEXT = 3;
        public static int CCPFTN = 4;
        public static int CCPHDD = 5;
        public static int CCPMCR = 6;
        public static int CCPATN = 7;
        public static int CCPEDN = 8;
        public static int CCPTXBX = 9;
        public static int CCPHDRTXBX = 10;
        public static int PNFBPCHPFIRST = 11;
        public static int PNCHPFIRST = 12;
        public static int CPNBTECHP = 13;
        public static int PNFBPPAPFIRST = 14;
        public static int PNPAPFIRST = 15;
        public static int CPNBTEPAP = 16;
        public static int PNFBPLVCFIRST = 17;
        public static int PNLVCFIRST = 18;
        public static int CPNBTELVC = 19;
        public static int FCISLANDFIRST = 20;
        public static int FCISLANDLIM = 21;

        int[] _longs;

        public FIBLongHandler(byte[] mainStream, int offset)
        {
            int longCount = LittleEndian.GetShort(mainStream, offset);
            offset += LittleEndianConstants.SHORT_SIZE;
            _longs = new int[longCount];

            for (int x = 0; x < longCount; x++)
            {
                _longs[x] = LittleEndian.GetInt(mainStream, offset + (x * LittleEndianConstants.INT_SIZE));
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
            offset += LittleEndianConstants.SHORT_SIZE;

            for (int x = 0; x < _longs.Length; x++)
            {
                LittleEndian.PutInt(mainStream, offset, _longs[x]);
                offset += LittleEndianConstants.INT_SIZE;
            }
        }

        internal int SizeInBytes()
        {
            return (_longs.Length * LittleEndianConstants.INT_SIZE) + LittleEndianConstants.SHORT_SIZE;
        }


    }
}

