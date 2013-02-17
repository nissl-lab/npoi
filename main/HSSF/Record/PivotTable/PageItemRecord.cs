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
     * SXPI - Page Item (0x00B6)<br/>
     * 
     * @author Patrick Cheng
     */
    public class PageItemRecord : StandardRecord
    {
        public const short sid = 0x00B6;
        private class FieldInfo
        {
            public const int ENCODED_SIZE = 6;
            /** Index to the View Item SXVI(0x00B2) record */
            private int _isxvi;
            /** Index to the {@link ViewFieldsRecord} SXVD(0x00B1) record */
            private int _isxvd;
            /** Object ID for the drop-down arrow */
            private int _idObj;

            public FieldInfo(RecordInputStream in1)
            {
                _isxvi = in1.ReadShort();
                _isxvd = in1.ReadShort();
                _idObj = in1.ReadShort();
            }

            internal void Serialize(ILittleEndianOutput out1)
            {
                out1.WriteShort(_isxvi);
                out1.WriteShort(_isxvd);
                out1.WriteShort(_idObj);
            }

            public void AppendDebugInfo(StringBuilder sb)
            {
                sb.Append('(');
                sb.Append("isxvi=").Append(HexDump.ShortToHex(_isxvi));
                sb.Append(" isxvd=").Append(HexDump.ShortToHex(_isxvd));
                sb.Append(" idObj=").Append(HexDump.ShortToHex(_idObj));
                sb.Append(')');
            }
        }

	    private FieldInfo[] _fieldInfos;

        public PageItemRecord(RecordInputStream in1)
        {
            int dataSize = in1.Remaining;
            if (dataSize % FieldInfo.ENCODED_SIZE != 0)
            {
                throw new RecordFormatException("Bad data size " + dataSize);
            }

            int nItems = dataSize / FieldInfo.ENCODED_SIZE;

            FieldInfo[] fis = new FieldInfo[nItems];
            for (int i = 0; i < fis.Length; i++)
            {
                fis[i] = new FieldInfo(in1);
            }
            _fieldInfos = fis;
        }


        public override void Serialize(ILittleEndianOutput out1)
        {
            for (int i = 0; i < _fieldInfos.Length; i++)
            {
                _fieldInfos[i].Serialize(out1);
            }
        }


        protected override int DataSize
        {
            get
            {
                return _fieldInfos.Length * FieldInfo.ENCODED_SIZE;
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
            StringBuilder sb = new StringBuilder();
            sb.Append("[SXPI]\n");
            for (int i = 0; i < _fieldInfos.Length; i++)
            {
                sb.Append("    item[").Append(i).Append("]=");
                _fieldInfos[i].AppendDebugInfo(sb);
                sb.Append('\n');
            }
            sb.Append("[/SXPI]\n");
            return sb.ToString();
        }
    }
}