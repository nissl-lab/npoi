/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */
namespace NPOI.HSSF.Record
{

    using System;
    using System.Text;
    using NPOI.Util;

    public class DrawingRecord : Record
    {
        public const short sid = 0xEC;

        private byte[] recordData;
        private byte[] contd;

        public DrawingRecord()
        {
        }

        public DrawingRecord(RecordInputStream in1)
        {
            recordData = in1.ReadRemainder();
        }

        public void ProcessContinueRecord(byte[] record)
        {
            //don't merge continue record with the drawing record, it must be Serialized Separately
            contd = record;
        }

        public override int Serialize(int offset, byte [] data)
        {
            if (recordData == null)
            {
                recordData = new byte[0];
            }
            LittleEndian.PutShort(data, 0 + offset, sid);
            LittleEndian.PutShort(data, 2 + offset, (short)(recordData.Length));
            if (recordData.Length > 0)
            {
                Array.Copy(recordData, 0, data, 4 + offset, recordData.Length);
            }
            return RecordSize;
        }

        public override int RecordSize
        {
            get
            {
                int retval = 4;

                if (recordData != null)
                {
                    retval += recordData.Length;
                }
                return retval;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public byte[] Data
        {
            get
            {
                if (contd != null)
                {
                    byte[] newBuffer = new byte[recordData.Length + contd.Length];
                    Array.Copy(recordData, 0, newBuffer, 0, recordData.Length);
                    Array.Copy(contd, 0, newBuffer, recordData.Length, contd.Length);
                    return newBuffer;
                }
                else
                {
                    return recordData;
                }
            }
            set 
            {
                this.recordData = value;
            }
        }

        public override Object Clone()
        {
            DrawingRecord rec = new DrawingRecord();

            if (recordData != null)
            {
                rec.recordData = new byte[recordData.Length];
                Array.Copy(recordData, 0, rec.recordData, 0, recordData.Length);
            }
            if (contd != null)
            {
                Array.Copy(contd, 0, rec.contd, 0, contd.Length);
                rec.contd = new byte[contd.Length];
            }

            return rec;
        }
    }
}