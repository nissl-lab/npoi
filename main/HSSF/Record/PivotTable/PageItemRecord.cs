/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.HSSF.Record.PivotTable
{

    using System;
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.Util;


    /**
     * SXPI - Page Item (0x00B6)<br/>
     * 
     * @author Patrick Cheng
     */
    public class PageItemRecord : StandardRecord
    {
        public static short sid = 0x00B6;

        private int isxvi;
        private int isxvd;
        private int idObj;

        public PageItemRecord(RecordInputStream in1)
        {
            isxvi = in1.ReadShort();
            isxvd = in1.ReadShort();
            idObj = in1.ReadShort();
        }


        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(isxvi);
            out1.WriteShort(isxvd);
            out1.WriteShort(idObj);
        }


        protected override int DataSize
        {
            get
            {
                return 2 + 2 + 2;
            }
        }


        public override short Sid
        {
            get
            {
                return sid;
            }
        }


        public override string ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[SXPI]\n");
            buffer.Append("    .isxvi      =").Append(HexDump.ShortToHex(isxvi)).Append('\n');
            buffer.Append("    .isxvd      =").Append(HexDump.ShortToHex(isxvd)).Append('\n');
            buffer.Append("    .idObj      =").Append(HexDump.ShortToHex(idObj)).Append('\n');

            buffer.Append("[/SXPI]\n");
            return buffer.ToString();
        }
    }
}