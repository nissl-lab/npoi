
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


namespace NPOI.HSSF.Record
{

    using System;
    using System.Text;
    using NPOI.Util;
    using System.Globalization;

    /**
     * Title:        Continue Record - Helper class used primarily for SST Records 
     * Description:  handles overflow for prior record in the input
     *               stream; content Is tailored to that prior record
     * @author Marc Johnson (mjohnson at apache dot org)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Csaba Nagy (ncsaba at yahoo dot com)
     * @version 2.0-pre
     */

    public class ContinueRecord : StandardRecord, ICloneable
    {
        public const short sid = 0x003C;
        private byte[] field_1_data;

        /**
         * default constructor
         */

        private ContinueRecord()
        {
        }

        public ContinueRecord(byte[] data)
        {
            field_1_data = data;
        }

        /**
         * Main constructor -- kinda dummy because we don't validate or fill fields
         *
         * @param in the RecordInputstream to Read the record from
         */

        public ContinueRecord(RecordInputStream in1)
        {
            field_1_data = in1.ReadRemainder();
        }
        protected override int DataSize
        {
            get
            {
                return field_1_data.Length;
            }
        }
        
        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.Write(field_1_data);
        }
        
        /*
         * USE ONLY within "ProcessContinue"
         */
        //public sealed override byte[] Serialize()
        //{
        //    byte[] retval = new byte[field_1_data.Length + 4];
        //    Serialize(0, retval);
        //    return retval;
        //}

        /**
         * Writes the full encoding of a Continue record without making an instance
         */
        [Obsolete]
        public static int Write(byte[] destBuf, int destOffset, byte? initialDataByte, byte[] srcData)
        {
            return Write(destBuf, destOffset, initialDataByte, srcData, 0, srcData.Length);
        }
        /**
         * @param initialDataByte (optional - often used for unicode flag). 
         * If supplied, this will be written before srcData
         * @return the total number of bytes written
         */
        [Obsolete]
        public static int Write(byte[] destBuf, int destOffset, byte? initialDataByte, byte[] srcData, int srcOffset, int len)
        {
            int totalLen = len + (initialDataByte == null ? 0 : 1);
            LittleEndian.PutUShort(destBuf, destOffset, sid);
            LittleEndian.PutUShort(destBuf, destOffset + 2, totalLen);
            int pos = destOffset + 4;
            if (initialDataByte != null)
            {
                LittleEndian.PutByte(destBuf, pos, Convert.ToByte(initialDataByte, CultureInfo.InvariantCulture));
                pos += 1;
            }
            Array.Copy(srcData, srcOffset, destBuf, pos, len);
            return 4 + totalLen;
        }
        //public override int Serialize(int offset, byte[] data)
        //{
        //    return Write(data, offset, null, field_1_data);
        //}

        /**
         * Get the data for continuation
         * @return byte array containing all of the continued data
         */

        public byte[] Data
        {
            get { return field_1_data; }
            set { field_1_data = value; }
        }


        /**
         * Debugging toString
         *
         * @return string representation
         */

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[CONTINUE RECORD]\n");
            buffer.Append("    .data        = ").Append(StringUtil.ToHexString(sid))
                .Append("\n");
            buffer.Append("[/CONTINUE RECORD]\n");
            return buffer.ToString();
        }

        public override short Sid
        {
            get { return sid; }
        }


        /**
         * Clone this record.
         */
        public override Object Clone()
        {
            ContinueRecord Clone = new ContinueRecord();
            Clone.Data = (field_1_data);
            return Clone;
        }

    }
}