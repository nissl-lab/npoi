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
    using NPOI.Util;
    using System.Text;

    /**
     * The UserSViewEnd record marks the end of the Settings for a custom view associated with the sheet
     *
     * @author Yegor Kozlov
     */
    public class UserSViewEnd : StandardRecord
    {

        public const short sid = 0x01AB;
        private byte[] _rawData;

        public UserSViewEnd(byte[] data)
        {
            _rawData = data;
        }

        /**
         * construct an UserSViewEnd record.  No fields are interpreted and the record will
         * be Serialized in its original form more or less
         * @param in the RecordInputstream to read the record from
         */
        public UserSViewEnd(RecordInputStream in1)
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

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("[").Append("USERSVIEWEND").Append("] (0x");
            sb.Append(StringUtil.ToHexString(sid).ToUpper() + ")\n");
            sb.Append("  rawData=").Append(HexDump.ToHex(_rawData)).Append("\n");
            sb.Append("[/").Append("USERSVIEWEND").Append("]\n");
            return sb.ToString();
        }

        //HACK: do a "cheat" Clone, see Record.java for more information
        public override Object Clone()
        {
            return CloneViaReserialise();
        }

    }
}
