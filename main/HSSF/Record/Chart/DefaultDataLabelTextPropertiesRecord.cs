
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


namespace NPOI.HSSF.Record.Chart
{

    using System.Text;
    using System;

    using NPOI.Util;


    public enum CategoryDataType:short
    {
        SHOW_LABELS_CHARACTERISTIC = 0,
        VALUE_AND_PERCENTAGE_CHARACTERISTIC =1,
        ALL_TEXT_CHARACTERISTIC0 =2,
        ALL_TEXT_CHARACTERISTIC1 =3
    }

    /**
     * The default data label text properties record identifies the text Charistics of the preceeding text record.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class DefaultDataLabelTextPropertiesRecord
       : StandardRecord
    {
        public const short sid = 0x1024;
        private short field_1_categoryDataType;

        public DefaultDataLabelTextPropertiesRecord()
        {

        }

        /**
         * Constructs a DefaultDataLabelTextProperties record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public DefaultDataLabelTextPropertiesRecord(RecordInputStream in1)
        {
            field_1_categoryDataType = in1.ReadShort();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[DEFAULTTEXT]\n");
            buffer.Append("    .categoryDataType     = ")
                .Append("0x").Append(HexDump.ToHex((short)CategoryDataType))
                .Append(" (").Append(CategoryDataType).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/DEFAULTTEXT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_categoryDataType);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            DefaultDataLabelTextPropertiesRecord rec = new DefaultDataLabelTextPropertiesRecord();

            rec.field_1_categoryDataType = field_1_categoryDataType;
            return rec;
        }




        /**
         * Get the category data type field for the DefaultDataLabelTextProperties record.
         *
         * @return  One of 
         *        CATEGORY_DATA_TYPE_SHOW_LABELS_CharISTIC
         *        CATEGORY_DATA_TYPE_VALUE_AND_PERCENTAGE_CharISTIC
         *        CATEGORY_DATA_TYPE_ALL_TEXT_CharISTIC
         */
        public CategoryDataType CategoryDataType
        {
            get
            {
                return (CategoryDataType)field_1_categoryDataType;
            }
            set 
            {
                this.field_1_categoryDataType = (short)value;
            }
        }


    }
}

