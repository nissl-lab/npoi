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
    using NPOI.Util;

    /**
     * STARTBLOCK - Chart Future Record Type Start Block (0x0852)<br/>
     * 
     * @author Patrick Cheng
     */
    public class StartBlockRecord : StandardRecord
    {
        public static short sid = 0x0852;

        private short rt;
        private short grbitFrt;

        public ObjectKind ObjectKind
        {
            get; set;
        }

        public short ObjectContext
        {
            get; set;
        }

        public short ObjectInstance1
        {
            get; set;
        }

        public short ObjectInstance2
        {
            get; set;
        }

        public StartBlockRecord()
        {
            rt = sid;
            grbitFrt = 0;
        }
        
        public StartBlockRecord(RecordInputStream in1)
        {
            rt = in1.ReadShort();
            grbitFrt = in1.ReadShort();
            ObjectKind = (ObjectKind) in1.ReadShort();
            ObjectContext = in1.ReadShort();
            ObjectInstance1 = in1.ReadShort();
            ObjectInstance2 = in1.ReadShort();
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
            out1.WriteShort((short) ObjectKind);
            out1.WriteShort(ObjectContext);
            out1.WriteShort(ObjectInstance1);
            out1.WriteShort(ObjectInstance2);
        }

        public override String ToString()
        {

            StringBuilder buffer = new StringBuilder();

            buffer.Append("[STARTBLOCK]\n");
            buffer.Append("    .rt              =").Append(HexDump.ShortToHex(rt)).Append('\n');
            buffer.Append("    .grbitFrt        =").Append(HexDump.ShortToHex(grbitFrt)).Append('\n');
            buffer.Append("    .iObjectKind     =").Append(HexDump.ShortToHex((short) ObjectKind)).Append('\n');
            buffer.Append("    .iObjectContext  =").Append(HexDump.ShortToHex(ObjectContext)).Append('\n');
            buffer.Append("    .iObjectInstance1=").Append(HexDump.ShortToHex(ObjectInstance1)).Append('\n');
            buffer.Append("    .iObjectInstance2=").Append(HexDump.ShortToHex(ObjectInstance2)).Append('\n');
            buffer.Append("[/STARTBLOCK]\n");
            return buffer.ToString();
        }
        public override object Clone()
        {
            StartBlockRecord record = new StartBlockRecord();
            record.rt = rt;
            record.grbitFrt = grbitFrt;
            record.ObjectKind = ObjectKind;
            record.ObjectContext = ObjectContext;
            record.ObjectInstance1 = ObjectInstance1;
            record.ObjectInstance2 = ObjectInstance2;
            return record;
        }
        
        public static StartBlockRecord CreateStartBlock(ObjectKind objectKind)
        {
            return CreateStartBlock(objectKind, 0, 0, 0);
        }

        public static StartBlockRecord CreateStartBlock(ObjectKind objectKind, short objectContext)
        {
            return CreateStartBlock(objectKind, objectContext, 0, 0);
        }

        public static StartBlockRecord CreateStartBlock(ObjectKind objectKind, short objectContext,
            short objectInstance1)
        {
            return CreateStartBlock(objectKind, objectContext, objectInstance1, 0);
        }

        public static StartBlockRecord CreateStartBlock(ObjectKind objectKind, short objectContext,
            short objectInstance1, short objectInstance2)
        {
            StartBlockRecord record = new StartBlockRecord();
            record.ObjectKind = objectKind;
            record.ObjectContext = objectContext;
            record.ObjectInstance1 = objectInstance1;
            record.ObjectInstance2 = objectInstance2;
            return record;
        }
    }
}