
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
    using System.Collections;
    using NPOI.Util;



    /**
     * Title: Beginning Of File
     * Description: Somewhat of a misnomer, its used for the beginning of a Set of
     *              records that have a particular pupose or subject.
     *              Used in sheets and workbooks.
     * REFERENCE:  PG 289 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class BOFRecord : StandardRecord
    {

        /**
         * for BIFF8 files the BOF is 0x809.  For earlier versions it was 0x09 or 0x(biffversion)09
         */

        public const short sid = 0x809;
        private int field_1_version;
        private int field_2_type;
        private int field_3_build;
        private int field_4_year;
        private int field_5_history;
        private int field_6_rversion;

        /**
         * suggested default (0x06 - BIFF8)
         */

        public const short VERSION = 0x06;

        /**
         * suggested default 0x10d3
         */

        public const short BUILD = 0x10d3;

        /**
         * suggested default  0x07CC (1996)
         */

        public const short BUILD_YEAR = 0x07CC;   // 1996

        /**
         * suggested default for a normal sheet (0x41)
         */

        public const short HISTORY_MASK = 0x41;
        public const short TYPE_WORKBOOK = 0x05;
        public const short TYPE_VB_MODULE = 0x06;
        public const short TYPE_WORKSHEET = 0x10;
        public const short TYPE_CHART = 0x20;
        public const short TYPE_EXCEL_4_MACRO = 0x40;
        public const short TYPE_WORKSPACE_FILE = 0x100;

        /**
         * Constructs an empty BOFRecord with no fields Set.
         */

        public BOFRecord()
        {
        }
        private BOFRecord(int type)
        {
            field_1_version = VERSION;
            field_2_type = type;
            field_3_build = BUILD;
            field_4_year = BUILD_YEAR;
            field_5_history = 0x01;
            field_6_rversion = VERSION;
        }

        public static BOFRecord CreateSheetBOF()
        {
            return new BOFRecord(TYPE_WORKSHEET);
        }

        /**
         * Constructs a BOFRecord and Sets its fields appropriately
         * @param in the RecordInputstream to Read the record from
         */

        public BOFRecord(RecordInputStream in1)
        {
            field_1_version = in1.ReadShort();
            field_2_type = in1.ReadShort();

            // Some external tools don't generate all of
            //  the remaining fields
            if (in1.Remaining >= 2)
            {
                field_3_build = in1.ReadShort();
            }
            if (in1.Remaining >= 2)
            {
                field_4_year = in1.ReadShort();
            }
            if (in1.Remaining >= 4)
            {
                field_5_history = in1.ReadInt();
            }
            if (in1.Remaining >= 4)
            {
                field_6_rversion = in1.ReadInt();
            }
        }

        /**
         * Version number - for BIFF8 should be 0x06
         * @see #VERSION
         * @param version version to be Set
         */

        public int Version
        {
            set { field_1_version = value; }
            get { return field_1_version; }
        }
        /**
         * Set the history bit mask (not very useful)
         * @see #HISTORY_MASK
         * @param bitmask bitmask to Set for the history
         */

        public int HistoryBitMask
        {
            set { field_5_history = value; }
            get { return field_5_history; }
        }

        /**
         * Set the minimum version required to Read this file
         *
         * @see #VERSION
         * @param version version to Set
         */

        public int RequiredVersion
        {
            set { field_6_rversion = value; }
            get { return field_6_rversion; }
        }

        /**
         * type of object this marks
         * @see #TYPE_WORKBOOK
         * @see #TYPE_VB_MODULE
         * @see #TYPE_WORKSHEET
         * @see #TYPE_CHART
         * @see #TYPE_EXCEL_4_MACRO
         * @see #TYPE_WORKSPACE_FILE
         * @return short type of object
         */

        public int Type
        {
            get { return field_2_type; }
            set { field_2_type = value; }
        }
        private String TypeName
        {
            get
            {
                switch (field_2_type)
                {
                    case TYPE_CHART: return "chart";
                    case TYPE_EXCEL_4_MACRO: return "excel 4 macro";
                    case TYPE_VB_MODULE: return "vb module";
                    case TYPE_WORKBOOK: return "workbook";
                    case TYPE_WORKSHEET: return "worksheet";
                    case TYPE_WORKSPACE_FILE: return "workspace file";
                }
                return "#error unknown type#";
            }
        }

        /**
         * Get the build that wrote this file
         * @see #BUILD
         * @return short build number of the generator of this file
         */

        public int Build
        {
            get { return field_3_build; }
            set { field_3_build = value; }
        }

        /**
         * Year of the build that wrote this file
         * @see #BUILD_YEAR
         * @return short build year of the generator of this file
         */

        public int BuildYear
        {
            get { return field_4_year; }
            set { field_4_year = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[BOF RECORD]\n");
            buffer.Append("    .version         = ")
                .Append(StringUtil.ToHexString(Version)).Append("\n");
            buffer.Append("    .type            = ")
                .Append(StringUtil.ToHexString(Type)).Append("\n");
            buffer.Append(" (").Append(TypeName).Append(")").Append("\n");
            buffer.Append("    .build           = ")
                .Append(StringUtil.ToHexString(Build)).Append("\n");
            buffer.Append("    .buildyear       = ").Append(BuildYear)
                .Append("\n");
            buffer.Append("    .history         = ")
                .Append(StringUtil.ToHexString(HistoryBitMask)).Append("\n");
            buffer.Append("    .requiredversion = ")
                .Append(StringUtil.ToHexString(RequiredVersion)).Append("\n");
            buffer.Append("[/BOF RECORD]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(Version);
            out1.WriteShort(Type);
            out1.WriteShort(Build);
            out1.WriteShort(BuildYear);
            out1.WriteInt(HistoryBitMask);
            out1.WriteInt(RequiredVersion);
        }

        protected override int DataSize
        {
            get { return 16; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override object Clone()
        {
            BOFRecord rec = new BOFRecord();
            rec.field_1_version = field_1_version;
            rec.field_2_type = field_2_type;
            rec.field_3_build = field_3_build;
            rec.field_4_year = field_4_year;
            rec.field_5_history = field_5_history;
            rec.field_6_rversion = field_6_rversion;
            return rec;
        }
    }
}