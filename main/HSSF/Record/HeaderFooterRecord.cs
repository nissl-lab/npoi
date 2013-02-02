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

namespace NPOI.HSSF.Record
{
    using NPOI.Util;

    using System;
    using System.Text;

    /**
     * The HEADERFOOTER record stores information Added in Office Excel 2007 for headers/footers.
     *
     * @author Yegor Kozlov
     */
    public class HeaderFooterRecord : StandardRecord
    {

        private static byte[] BLANK_GUID = new byte[16];

        public const short sid = 0x089C;
        private byte[] _rawData;

        public HeaderFooterRecord(byte[] data)
        {
            _rawData = data;
        }

        /**
         * construct a HeaderFooterRecord record.  No fields are interpreted and the record will
         * be Serialized in its original form more or less
         * @param in the RecordInputstream to read the record from
         */
        public HeaderFooterRecord(RecordInputStream in1)
        {
            _rawData = in1.ReadRemainder();
        }

        /**
         * spit the record out AS IS. no interpretation or identification
         */
        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.Write(_rawData);
        }

        protected override int DataSize
        {
            get
            {
                return _rawData.Length;
            }
        }

        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        /**
         * If this header belongs to a specific sheet view , the sheet view?s GUID will be saved here.
         * 
         * If it is zero, it means the current sheet. Otherwise, this field MUST match the guid field
         * of the preceding {@link UserSViewBegin} record.
         *
         * @return the sheet view's GUID
         */
        public byte[] Guid
        {
            get
            {
                byte[] guid = new byte[16];
                Array.Copy(_rawData, 12, guid, 0, guid.Length);
                return guid;
            }
        }

        /**
         * @return whether this record belongs to the current sheet 
         */
        public bool IsCurrentSheet
        {
            get
            {
                return Arrays.Equals(Guid, BLANK_GUID);
            }
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("[").Append("HEADERFOOTER").Append("] (0x");
            sb.Append(StringUtil.ToHexString(sid).ToUpper() + ")\n");
            sb.Append("  rawData=").Append(HexDump.ToHex(_rawData)).Append("\n");
            sb.Append("[/").Append("HEADERFOOTER").Append("]\n");
            return sb.ToString();
        }

        //HACK: do a "cheat" Clone, see Record.java for more information
        public override Object Clone()
        {
            return CloneViaReserialise();
        }
    }
}


