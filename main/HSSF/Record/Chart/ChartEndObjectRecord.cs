/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HSSF.Record.Chart
{

    using System;
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    using NPOI.Util.IO;

    /**
     * ENDOBJECT - Chart Future Record Type End Object (0x0855)<br/>
     * 
     * @author Patrick Cheng
     */
    public class ChartEndObjectRecord : StandardRecord
    {
        public static short sid = 0x0855;

        private short rt;
        private short grbitFrt;
        private short iObjectKind;
        private byte[] unused;

        public ChartEndObjectRecord(RecordInputStream in1)
        {
            rt = in1.ReadShort();
            grbitFrt = in1.ReadShort();
            iObjectKind = in1.ReadShort();

            unused = new byte[6];
            in1.ReadFully(unused);
        }


        protected override int DataSize
        {
            get
            {
                return 2 + 2 + 2 + 6;
            }
        }

        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        public override void Serialize(LittleEndianOutput out1)
        {
            out1.WriteShort(rt);
            out1.WriteShort(grbitFrt);
            out1.WriteShort(iObjectKind);
            // 6 bytes unused
            out1.Write(unused);
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[ENDOBJECT]\n");
            buffer.Append("    .rt         =").Append(HexDump.ShortToHex(rt)).Append('\n');
            buffer.Append("    .grbitFrt   =").Append(HexDump.ShortToHex(grbitFrt)).Append('\n');
            buffer.Append("    .iObjectKind=").Append(HexDump.ShortToHex(iObjectKind)).Append('\n');
            buffer.Append("    .unused     =").Append(HexDump.ToHex(unused)).Append('\n');
            buffer.Append("[/ENDOBJECT]\n");
            return buffer.ToString();
        }
    }
}
