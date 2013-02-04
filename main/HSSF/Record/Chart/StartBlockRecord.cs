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
        public static short ObjectKind_AxisGroup = 0;
        public static short ObjectKind_AttachedLabelRecord = 0x2;
        public static short ObjectKind_Axis = 0x4;
        public static short ObjectKind_ChartGroup = 0x5;
        public static short ObjectKind_DatRecord = 0x6;
        public static short ObjectKind_Frame = 0x7;
        public static short ObjectKind_Legend = 0x9;
        public static short ObjectKind_LegendException = 0xA;
        public static short ObjectKind_Series = 0xC;
        public static short ObjectKind_Sheet = 0xD;
        public static short ObjectKind_DataFormatRecord = 0xE;
        public static short ObjectKind_DropBarRecord = 0xF;

        public static short sid = 0x0852;

        private short rt;
        private short grbitFrt;
        private short iObjectKind;

        public short ObjectKind
        {
            get { return iObjectKind; }
            set { iObjectKind = value; }
        }
        private short iObjectContext;

        public short ObjectContext
        {
            get { return iObjectContext; }
            set { iObjectContext = value; }
        }
        private short iObjectInstance1;

        public short ObjectInstance1
        {
            get { return iObjectInstance1; }
            set { iObjectInstance1 = value; }
        }
        private short iObjectInstance2;

        public short ObjectInstance2
        {
            get { return iObjectInstance2; }
            set { iObjectInstance2 = value; }
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
            StartBlockRecord record = new StartBlockRecord();
            record.rt = rt;
            record.grbitFrt = grbitFrt;
            record.iObjectKind = iObjectKind;
            record.iObjectContext = iObjectContext;
            record.iObjectInstance1 = iObjectInstance1;
            record.iObjectInstance2 = iObjectInstance2;
            return record;
        }
        
        public static StartBlockRecord CreateStartBlock(short objectKind)
        {
            return CreateStartBlock(objectKind, 0, 0, 0);
        }

        public static StartBlockRecord CreateStartBlock(short objectKind, short objectContext)
        {
            return CreateStartBlock(objectKind, objectContext, 0, 0);
        }

        public static StartBlockRecord CreateStartBlock(short objectKind, short objectContext,
            short objectInstance1)
        {
            return CreateStartBlock(objectKind, objectContext, objectInstance1, 0);
        }

        public static StartBlockRecord CreateStartBlock(short objectKind, short objectContext,
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