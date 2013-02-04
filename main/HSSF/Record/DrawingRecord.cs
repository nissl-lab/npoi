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
    using NPOI.Util;
    /**
     * DrawingRecord (0x00EC)<p/>
     *
     */
    public class DrawingRecord : StandardRecord
    {
        public const short sid = 0xEC;
        private static byte[] EMPTY_BYTE_ARRAY = { };
        private byte[] recordData;
        private byte[] contd;

        public DrawingRecord()
        {
            recordData = EMPTY_BYTE_ARRAY;
        }

        public DrawingRecord(RecordInputStream in1)
        {
            recordData = in1.ReadRemainder();
        }
        [Obsolete]
        public void ProcessContinueRecord(byte[] record)
        {
            //don't merge continue record with the drawing record, it must be Serialized Separately
            contd = record;
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.Write(recordData);
        }

        protected override int DataSize
        {
            get
            {
                return recordData.Length;
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
                //if (contd != null)
                //{
                //    byte[] newBuffer = new byte[recordData.Length + contd.Length];
                //    Array.Copy(recordData, 0, newBuffer, 0, recordData.Length);
                //    Array.Copy(contd, 0, newBuffer, recordData.Length, contd.Length);
                //    return newBuffer;
                //}
                return recordData;
            }
            set 
            {
                if (value == null)
                {
                    throw new ArgumentException("data must not be null");
                }
                this.recordData = value;
            }
        }
        /**
         * Cloning of drawing records must be executed through HSSFPatriarch, because all id's must be changed
         * @return cloned drawing records
         */
        public override Object Clone()
        {
            DrawingRecord rec = new DrawingRecord();

            rec.recordData = (byte[])recordData.Clone();// new byte[recordData.Length];
            if (contd != null)
            {
                rec.contd = (byte[])contd.Clone();
            }

            return rec;
        }

        public override String ToString()
        {
            return "DrawingRecord[" + recordData.Length + "]";
        }
    }
}