
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

    using NPOI.Util;
    using System;
    using System.Text;

    /**
     * Title:        Default Row Height Record
     * Description:  Row height for rows with Undefined or not explicitly defined
     *               heights.
     * REFERENCE:  PG 301 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class DefaultRowHeightRecord
       : StandardRecord
    {
        public const short sid = 0x225;
        private short field_1_option_flags;
        private short field_2_row_height;
        /**
     * The default row height for empty rows is 255 twips (255 / 20 == 12.75 points)
     */
        public const short DEFAULT_ROW_HEIGHT = 0xFF;
        //BitField isHeightChanged = BitFieldFactory.GetInstance(0x01);
        //BitField isZeroHeight = BitFieldFactory.GetInstance(0x02);
        //BitField isThickTopBorder = BitFieldFactory.GetInstance(0x04);
        //BitField isThickBottomBorder = BitFieldFactory.GetInstance(0x08);

        public DefaultRowHeightRecord()
        {
            field_1_option_flags = 0x0000;
            field_2_row_height = DEFAULT_ROW_HEIGHT;
        }

       /// <summary>
       /// Constructs a DefaultRowHeight record and Sets its fields appropriately.
       /// </summary>
       /// <param name="in1">the RecordInputstream to Read the record from</param>
        public DefaultRowHeightRecord(RecordInputStream in1)
        {
            field_1_option_flags = in1.ReadShort();
            field_2_row_height = in1.ReadShort();
        }
        internal short OptionFlags
        {
            get { return field_1_option_flags; }
            set { field_1_option_flags = value; }
        }
        ///// <summary>
        ///// A bit that specifies whether the default settings for the row height have been changed.
        ///// </summary>
        //public bool IsDefaultHeightChanged
        //{
        //    get { return isHeightChanged.IsSet(field_1_option_flags); }
        //    set { field_1_option_flags = isHeightChanged.SetShortBoolean(field_1_option_flags,value); }
        //}
        ///// <summary>
        ///// A bit that specifies whether empty rows have a height of zero.
        ///// </summary>
        //public bool IsZeroHeight
        //{
        //    get { return isZeroHeight.IsSet(field_1_option_flags); }
        //    set { field_1_option_flags=isZeroHeight.SetShortBoolean(field_1_option_flags,value); }
        //}

        //public bool IsThickTopBorder
        //{
        //    get { return isThickTopBorder.IsSet(field_1_option_flags); }
        //    set { field_1_option_flags = isThickTopBorder.SetShortBoolean(field_1_option_flags, value); }        
        
        //}
        //public bool IsThickBottomBorder
        //{
        //    get { return isThickBottomBorder.IsSet(field_1_option_flags); }
        //    set { field_1_option_flags = isThickBottomBorder.SetShortBoolean(field_1_option_flags, value); }
        //}

        /// <summary>
        /// Get the default row height
        /// </summary>
        public short RowHeight
        {
            get { return field_2_row_height; }
            set {
                field_2_row_height = value; 
            }
        }

        
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[DEFAULTROWHEIGHT]\n");
            buffer.Append("    .optionflags    = ")
                .Append(StringUtil.ToHexString(OptionFlags)).Append("\n");
            buffer.Append("    .rowheight      = ")
                .Append(StringUtil.ToHexString(RowHeight)).Append("\n");
            buffer.Append("[/DEFAULTROWHEIGHT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(OptionFlags);
            out1.WriteShort(RowHeight);
        }

        protected override int DataSize
        {
            get
            {
                return 4;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            DefaultRowHeightRecord rec = new DefaultRowHeightRecord();
            rec.field_1_option_flags = field_1_option_flags;
            rec.field_2_row_height = field_2_row_height;
            return rec;
        }
    }
}
