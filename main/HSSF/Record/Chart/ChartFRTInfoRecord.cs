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


    /*
     * CHARTFRTINFO - Chart Future Record Type Info (0x0850)<br/>
     * 
     * @author Patrick Cheng
     */
    //
    /// <summary>
    /// The ChartFrtInfo record specifies the versions of the application that originally created and last saved the file.
    /// </summary>
    public class ChartFRTInfoRecord : StandardRecord
    {
        public const short sid = 0x850;

        private short rt;
        private short grbitFrt;
        private byte verOriginator;
        private byte verWriter;
        private CFRTID[] rgCFRTID;

        private class CFRTID : ICloneable
        {
            public const int ENCODED_SIZE = 4;
            private int rtFirst;
            private int rtLast;

            public CFRTID(int first, int last)
            {
                rtFirst = first;
                rtLast = last;
            }
            public CFRTID(RecordInputStream in1)
            {
                rtFirst = in1.ReadShort();
                rtLast = in1.ReadShort();
            }

            public void Serialize(ILittleEndianOutput out1)
            {
                out1.WriteShort(rtFirst);
                out1.WriteShort(rtLast);
            }

            #region ICloneable ��Ա

            public object Clone()
            {
                return new CFRTID(this.rtFirst, this.rtLast);
            }

            #endregion
        }

        public ChartFRTInfoRecord()
        {
            rt = 0x850;
            grbitFrt = 0;
            //created by excel 2003
            verOriginator = (byte)0xA;
            //writen by excel 2003
            verWriter = (byte)0xA;
            rgCFRTID = new CFRTID[3];
            rgCFRTID[0] = new CFRTID(0x0850, 0x085A);
            rgCFRTID[1] = new CFRTID(0x0861, 0x0861);
            rgCFRTID[2] = new CFRTID(0x086A, 0x086B);
        }

        public ChartFRTInfoRecord(RecordInputStream in1)
        {
            rt = in1.ReadShort();
            grbitFrt = in1.ReadShort();
            verOriginator = (byte)in1.ReadByte();
            verWriter = (byte)in1.ReadByte();
            int cCFRTID = in1.ReadShort();

            rgCFRTID = new CFRTID[cCFRTID];
            for (int i = 0; i < cCFRTID; i++)
            {
                rgCFRTID[i] = new CFRTID(in1);
            }
        }


        protected override int DataSize
        {
            get
            {
                return 2 + 2 + 1 + 1 + 2 + rgCFRTID.Length * CFRTID.ENCODED_SIZE;
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
            out1.WriteByte(verOriginator);
            out1.WriteByte(verWriter);
            int nCFRTIDs = rgCFRTID.Length;
            out1.WriteShort(nCFRTIDs);

            for (int i = 0; i < nCFRTIDs; i++)
            {
                rgCFRTID[i].Serialize(out1);
            }
        }

        public override object Clone()
        {
            ChartFRTInfoRecord record = new ChartFRTInfoRecord();

            record.grbitFrt = this.grbitFrt;
            record.rgCFRTID = new CFRTID[this.rgCFRTID.Length];
            record.rt = this.rt;
            record.verOriginator = this.verOriginator;
            record.verWriter = verWriter;

            for (int i = 0; i < this.rgCFRTID.Length; i++)
                record.rgCFRTID[i] = (CFRTID)this.rgCFRTID[i].Clone();

            return record;
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[CHARTFRTINFO]\n");
            buffer.Append("    .rt           =").Append(HexDump.ShortToHex(rt)).Append('\n');
            buffer.Append("    .grbitFrt     =").Append(HexDump.ShortToHex(grbitFrt)).Append('\n');
            buffer.Append("    .verOriginator=").Append(HexDump.ByteToHex(verOriginator)).Append('\n');
            buffer.Append("    .verWriter    =").Append(HexDump.ByteToHex(verOriginator)).Append('\n');
            buffer.Append("    .nCFRTIDs     =").Append(HexDump.ShortToHex(rgCFRTID.Length)).Append('\n');
            buffer.Append("[/CHARTFRTINFO]\n");
            return buffer.ToString();
        }
    }
}