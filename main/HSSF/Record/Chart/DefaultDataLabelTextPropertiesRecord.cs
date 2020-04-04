/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI.HSSF.Record.Chart
{
    using System;
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.Util;

    /**
     * The default data label text properties record identifies the text characteristics of the preceding text record.<p/>
     * 
     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class DefaultDataLabelTextPropertiesRecord : StandardRecord, ICloneable
    {
        public static short sid = 0x1024;
        private short field_1_categoryDataType;
        public static short CATEGORY_DATA_TYPE_SHOW_LABELS_CHARACTERISTIC = 0;
        public static short CATEGORY_DATA_TYPE_VALUE_AND_PERCENTAGE_CHARACTERISTIC = 1;
        public static short CATEGORY_DATA_TYPE_ALL_TEXT_CHARACTERISTIC = 2;


        public DefaultDataLabelTextPropertiesRecord()
        {

        }

        public DefaultDataLabelTextPropertiesRecord(RecordInputStream in1)
        {
            field_1_categoryDataType = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[DEFAULTTEXT]\n");
            buffer.Append("    .categoryDataType     = ")
                .Append("0x").Append(HexDump.ToHex(CategoryDataType))
                .Append(" (").Append(CategoryDataType).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/DEFAULTTEXT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_categoryDataType);
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


        public override object Clone()
        {
            DefaultDataLabelTextPropertiesRecord rec = new DefaultDataLabelTextPropertiesRecord();

            rec.field_1_categoryDataType = field_1_categoryDataType;
            return rec;
        }




        /**
         * Get the category data type field for the DefaultDataLabelTextProperties record.
         *
         * @return  One of 
         *        CATEGORY_DATA_TYPE_SHOW_LABELS_CHARACTERISTIC
         *        CATEGORY_DATA_TYPE_VALUE_AND_PERCENTAGE_CHARACTERISTIC
         *        CATEGORY_DATA_TYPE_ALL_TEXT_CHARACTERISTIC
         */
        public short CategoryDataType
        {
            get
            {
                return field_1_categoryDataType;
            }
            set
            {
                this.field_1_categoryDataType = value;
            }   
        }
    }
}