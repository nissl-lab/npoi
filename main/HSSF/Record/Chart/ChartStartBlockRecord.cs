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
    using System.IO;
    using System.Text;
    using NPOI.Util;


    /**
     * STARTBLOCK - Chart Future Record Type Start Block (0x0852)<br/>
     * 
     * @author Patrick Cheng
     */
    public class ChartStartBlockRecord : StandardRecord
    {
        public static short sid = 0x0852;

        private short rt;
        private short grbitFrt;
        private short iObjectKind;
        private short iObjectContext;
        private short iObjectInstance1;
        private short iObjectInstance2;

        public ChartStartBlockRecord()
        { }
        
        public ChartStartBlockRecord(RecordInputStream in1)
        {
            rt = in1.ReadShort();
            grbitFrt = in1.ReadShort();
            iObjectKind = in1.ReadShort();
            iObjectContext = in1.ReadShort();
            iObjectInstance1 = in1.ReadShort();
            iObjectInstance2 = in1.ReadShort();
        }


        protected override int DataSize
        {
            get
            {
                return 2 + 2 + 2 + 2 + 2 + 2;
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
            out1.WriteShort(iObjectKind);
            out1.WriteShort(iObjectContext);
            out1.WriteShort(iObjectInstance1);
            out1.WriteShort(iObjectInstance2);
        }

        public override String ToString()
        {

            StringBuilder buffer = new StringBuilder();

            buffer.Append("[STARTBLOCK]\n");
            buffer.Append("    .rt              =").Append(HexDump.ShortToHex(rt)).Append('\n');
            buffer.Append("    .grbitFrt        =").Append(HexDump.ShortToHex(grbitFrt)).Append('\n');
            buffer.Append("    .iObjectKind     =").Append(HexDump.ShortToHex(iObjectKind)).Append('\n');
            buffer.Append("    .iObjectContext  =").Append(HexDump.ShortToHex(iObjectContext)).Append('\n');
            buffer.Append("    .iObjectInstance1=").Append(HexDump.ShortToHex(iObjectInstance1)).Append('\n');
            buffer.Append("    .iObjectInstance2=").Append(HexDump.ShortToHex(iObjectInstance2)).Append('\n');
            buffer.Append("[/STARTBLOCK]\n");
            return buffer.ToString();
        }
        public override object Clone()
        {
            ChartStartBlockRecord record = new ChartStartBlockRecord();
            record.rt = rt;
            record.grbitFrt = grbitFrt;
            record.iObjectKind = iObjectKind;
            record.iObjectContext = iObjectContext;
            record.iObjectInstance1 = iObjectInstance1;
            record.iObjectInstance2 = iObjectInstance2;
            return record;
        }

    }
}