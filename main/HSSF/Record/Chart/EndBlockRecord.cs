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
     * ENDBLOCK - Chart Future Record Type End Block (0x0853)<br/>
     * 
     * @author Patrick Cheng
     */
    public class EndBlockRecord : StandardRecord
    {
        public const short sid = 0x0853;

        private short rt;
        private short grbitFrt;

        public ObjectKind ObjectKind
        {
            get; set;
        }
        private byte[] unused;

        public EndBlockRecord()
        {
            rt = sid;
            grbitFrt = 0;
        }

        public EndBlockRecord(RecordInputStream in1)
        {
            rt = in1.ReadShort();
            grbitFrt = in1.ReadShort();
            ObjectKind = (ObjectKind) in1.ReadShort();
            // Often, but not always has 6 unused bytes at the end
		    if(in1.Available() == 0) {
			    unused = new byte[0];
		    } else {
			    unused = new byte[6];
			    in1.ReadFully(unused);
		    }

        }


        protected override int DataSize
        {
            get
            {
                return 2 + 2 + 2 + unused.Length;
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
            out1.WriteShort((short) ObjectKind);
            // 6 bytes unused
            out1.Write(unused);
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[ENDBLOCK]\n");
            buffer.Append("    .rt         =").Append(HexDump.ShortToHex(rt)).Append('\n');
            buffer.Append("    .grbitFrt   =").Append(HexDump.ShortToHex(grbitFrt)).Append('\n');
            buffer.Append("    .iObjectKind=").Append(HexDump.ShortToHex((short) ObjectKind)).Append('\n');
            //buffer.Append("    .unused     =").Append(HexDump.ToHex(unused)).Append('\n');
            buffer.Append("[/ENDBLOCK]\n");
            return buffer.ToString();
        }

        public EndBlockRecord clone()
        {
            EndBlockRecord record = new EndBlockRecord();
            record.rt = rt;
            record.grbitFrt = grbitFrt;
            record.ObjectKind = ObjectKind;
            record.unused = (byte[])unused.Clone();
            return record;
        }

        public static EndBlockRecord CreateEndBlock(ObjectKind objectKind)
        {
            EndBlockRecord record = new EndBlockRecord();
            record.ObjectKind = objectKind;
            return record;
        }
    }
}