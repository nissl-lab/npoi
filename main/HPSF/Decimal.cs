/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */
using NPOI.Util;

namespace NPOI.HPSF
{
    public class Decimal
    {
        public const int SIZE = 16;

        private short field_1_wReserved;
        private byte field_2_scale;
        private byte field_3_sign;
        private int field_4_hi32;
        private long field_5_lo64;

        public Decimal(byte[] data, int startOffset)
        {
            int offset = startOffset;

            field_1_wReserved = LittleEndian.GetShort(data, offset);
            offset += LittleEndian.SHORT_SIZE;

            field_2_scale = data[offset];
            offset += LittleEndian.BYTE_SIZE;

            field_3_sign = data[offset];
            offset += LittleEndian.BYTE_SIZE;

            field_4_hi32 = LittleEndian.GetInt(data, offset);
            offset += LittleEndian.INT_SIZE;

            field_5_lo64 = LittleEndian.GetLong(data, offset);
            offset += LittleEndian.LONG_SIZE;
        }
    }
}