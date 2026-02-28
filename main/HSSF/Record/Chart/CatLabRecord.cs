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


    /**
     * CATLAB - Category Labels (0x0856)<br/>
     * 
     * @author Patrick Cheng
     */
    public class CatLabRecord : StandardRecord
    {
        public const short sid = 0x0856;

        private short rt;
        private short grbitFrt;
        private short wOffset;
        private short at;
        private short grbit;
        private short? unused;

        public CatLabRecord(RecordInputStream in1)
        {
            rt = in1.ReadShort();
            grbitFrt = in1.ReadShort();
            wOffset = in1.ReadShort();
            at = in1.ReadShort();
            grbit = in1.ReadShort();
            // Often, but not always has an unused short at the end
            if (in1.Available() == 0)
            {
                unused = null;
            }
            else
            {
                unused = in1.ReadShort();
            }

        }


        protected override int DataSize
        {
            get
            {
                return 2 + 2 + 2 + 2 + 2 + (unused == null ? 0 : 2);
            }
        }

        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        public override void Serialize(ILittleEndianOutput out1)
        {

            out1.WriteShort(rt);
            out1.WriteShort(grbitFrt);
            out1.WriteShort(wOffset);
            out1.WriteShort(at);
            out1.WriteShort(grbit);
            if (unused != null)
                out1.WriteShort((short)unused);
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[CATLAB]\n");
            buffer.Append("    .rt      =").Append(HexDump.ShortToHex(rt)).Append('\n');
            buffer.Append("    .grbitFrt=").Append(HexDump.ShortToHex(grbitFrt)).Append('\n');
            buffer.Append("    .wOffset =").Append(HexDump.ShortToHex(wOffset)).Append('\n');
            buffer.Append("    .at      =").Append(HexDump.ShortToHex(at)).Append('\n');
            buffer.Append("    .grbit   =").Append(HexDump.ShortToHex(grbit)).Append('\n');
            buffer.Append("    .unused  =").Append(HexDump.ShortToHex((short)unused)).Append('\n');

            buffer.Append("[/CATLAB]\n");
            return buffer.ToString();
        }
    }
}