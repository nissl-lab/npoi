
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
    using NPOI.HSSF.Record.Cont;

    /**
     * Supports the STRING record structure.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [Serializable]
    public class StringRecord : ContinuableRecord
    {
        public const short sid = 0x207;
        private bool _is16bitUnicode;
        //int field_1_string_length;
        //private String field_3_string;
        private String _text;

        public StringRecord()
        {
        }

        /**
         * Constructs a String record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */
        public StringRecord(RecordInputStream in1)
        {
            int field_1_string_length = in1.ReadShort();
            _is16bitUnicode = in1.ReadByte() != 0x00;

            if (_is16bitUnicode)
            {
                _text = in1.ReadUnicodeLEString(field_1_string_length);
            }
            else
            {
                _text = in1.ReadCompressedUnicode(field_1_string_length);
            }
        }

        /**
         * called by the class that Is responsible for writing this sucker.
         * Subclasses should implement this so that their data Is passed back in a
         * byte array.
         *
         * @param offset to begin writing at
         * @param data byte array containing instance data
         * @return number of bytes written
         */
        protected override void Serialize(ContinuableRecordOutput out1)
        {
            out1.WriteShort(this._text.Length);
            out1.WriteStringData(this._text);
        }
        //public void ProcessContinueRecord(byte[] data)
        //{

        //    if (!_is16bitUnicode)
        //    {
        //        field_3_string += StringUtil.GetFromUnicodeLE(data, 0, field_1_string_length - field_3_string.Length);
        //    }
        //    else
        //    {
        //        field_3_string += StringUtil.GetFromCompressedUnicode(data, 0, field_1_string_length - field_3_string.Length);
        //    }
        //} 
        /**
         * return the non static version of the id for this record.
         */
        public override short Sid
        {
            get { return sid; }
        }

        /**
         * @return The string represented by this record.
         */
        public String String
        {
            get { return _text; }
            set
            {
                this._text = value;
                this._is16bitUnicode = StringUtil.HasMultibyte(value);
            }
        }


        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[STRING]\n");
            buffer.Append("    .string            = ")
                .Append(_text).Append("\n");
            buffer.Append("[/STRING]\n");
            return buffer.ToString();
        }

        public override Object Clone()
        {
            StringRecord rec = new StringRecord();
            rec._is16bitUnicode = this._is16bitUnicode;
            rec._text = this._text;
            return rec;
        }

    }
}