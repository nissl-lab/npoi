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
    using System;
using System.Text;
using NPOI.Util;


/**
 * The UserSViewBegin record specifies Settings for a custom view associated with the sheet.
 * This record also marks the start of custom view records, which save custom view Settings.
 * Records between {@link UserSViewBegin} and {@link UserSViewEnd} contain Settings for the custom view,
 * not Settings for the sheet itself.
 *
 * @author Yegor Kozlov
 */
    public class UserSViewBegin : StandardRecord
    {

        public const short sid = 0x01AA;
        private byte[] _rawData;

        public UserSViewBegin(byte[] data)
        {
            _rawData = data;
        }

        /**
         * construct an UserSViewBegin record.  No fields are interpreted and the record will
         * be Serialized in its original form more or less
         * @param in the RecordInputstream to read the record from
         */
        public UserSViewBegin(RecordInputStream in1)
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
         * @return Globally unique identifier for the custom view
         */
        public byte[] Guid
        {
            get
            {
                byte[] guid = new byte[16];
                Array.Copy(_rawData, 0, guid, 0, guid.Length);
                return guid;
            }
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("[").Append("USERSVIEWBEGIN").Append("] (0x");
            sb.Append(StringUtil.ToHexString(sid).ToUpper() + ")\n");
            sb.Append("  rawData=").Append(HexDump.ToHex(_rawData)).Append("\n");
            sb.Append("[/").Append("USERSVIEWBEGIN").Append("]\n");
            return sb.ToString();
        }

        //HACK: do a "cheat" Clone, see Record.java for more information
        public override Object Clone()
        {
            return CloneViaReserialise();
        }
    }
}

