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
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.Util;



    /**
     * SXVS - View Source (0x00E3)<br/>
     * 
     * @author Patrick Cheng
     */
    public class ViewSourceRecord : StandardRecord
    {
        public const short sid = 0x00E3;

        private int vs;

        public ViewSourceRecord(RecordInputStream in1)
        {
            vs = in1.ReadShort();
        }


        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(vs);
        }


        protected override int DataSize
        {
            get
            {
                return 2;
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

            buffer.Append("[SXVS]\n");
            buffer.Append("    .vs      =").Append(HexDump.ShortToHex(vs)).Append('\n');

            buffer.Append("[/SXVS]\n");
            return buffer.ToString();
        }
    }
}