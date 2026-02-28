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

    using NPOI.HSSF.Record;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Formula;
    using NPOI.Util;
    using System.Text;

    /**
     * Describes a linked data record.  This record refers to the series data or text.<p/>
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class LinkedDataRecord : StandardRecord, ICloneable
    {
        public static short sid = 0x1051;

        private static BitField customNumberFormat = BitFieldFactory.GetInstance(0x1);

        private byte field_1_linkType;
        public static byte LINK_TYPE_TITLE_OR_TEXT = 0;
        public static byte LINK_TYPE_VALUES = 1;
        public static byte LINK_TYPE_CATEGORIES = 2;
        public static byte LINK_TYPE_SECONDARY_CATEGORIES = 3;
        private byte field_2_referenceType;
        public static byte REFERENCE_TYPE_DEFAULT_CATEGORIES = 0;
        public static byte REFERENCE_TYPE_DIRECT = 1;
        public static byte REFERENCE_TYPE_WORKSHEET = 2;
        public static byte REFERENCE_TYPE_NOT_USED = 3;
        public static byte REFERENCE_TYPE_ERROR_REPORTED = 4;
        private short field_3_options;
        private short field_4_indexNumberFmtRecord;
        private Formula field_5_formulaOfLink;


        public LinkedDataRecord()
        {

        }

        public LinkedDataRecord(RecordInputStream in1)
        {
            field_1_linkType = (byte)in1.ReadByte();
            field_2_referenceType = (byte)in1.ReadByte();
            field_3_options = in1.ReadShort();
            field_4_indexNumberFmtRecord = in1.ReadShort();
            int encodedTokenLen = in1.ReadUShort();
            field_5_formulaOfLink = Formula.Read(encodedTokenLen, in1);
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[AI]\n");
            buffer.Append("    .linkType             = ").Append(HexDump.ByteToHex(LinkType)).Append('\n');
            buffer.Append("    .referenceType        = ").Append(HexDump.ByteToHex(ReferenceType)).Append('\n');
            buffer.Append("    .options              = ").Append(HexDump.ShortToHex(Options)).Append('\n');
            buffer.Append("    .customNumberFormat   = ").Append(IsCustomNumberFormat).Append('\n');
            buffer.Append("    .indexNumberFmtRecord = ").Append(HexDump.ShortToHex(IndexNumberFmtRecord)).Append('\n');
            buffer.Append("    .FormulaOfLink        = ").Append('\n');
            Ptg[] ptgs = field_5_formulaOfLink.Tokens;
            for (int i = 0; i < ptgs.Length; i++)
            {
                Ptg ptg = ptgs[i];
                buffer.Append(ptg.ToString()).Append(ptg.RVAType).Append('\n');
            }

            buffer.Append("[/AI]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteByte(field_1_linkType);
            out1.WriteByte(field_2_referenceType);
            out1.WriteShort(field_3_options);
            out1.WriteShort(field_4_indexNumberFmtRecord);
            field_5_formulaOfLink.Serialize(out1);
        }

        protected override int DataSize
        {
            get
            {
                return 1 + 1 + 2 + 2 + field_5_formulaOfLink.EncodedSize;
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
            LinkedDataRecord rec = new LinkedDataRecord();

            rec.field_1_linkType = field_1_linkType;
            rec.field_2_referenceType = field_2_referenceType;
            rec.field_3_options = field_3_options;
            rec.field_4_indexNumberFmtRecord = field_4_indexNumberFmtRecord;
            rec.field_5_formulaOfLink = field_5_formulaOfLink.Copy();
            return rec;
        }




        /**
         * Get the link type field for the LinkedData record.
         *
         * @return  One of
         *        LINK_TYPE_TITLE_OR_TEXT
         *        LINK_TYPE_VALUES
         *        LINK_TYPE_CATEGORIES
         */
        public byte LinkType
        {
            get
            {
                return field_1_linkType;
            }
            set
            {
                this.field_1_linkType = value;
            }
        }

        /**
         * Get the reference type field for the LinkedData record.
         *
         * @return  One of
         *        REFERENCE_TYPE_DEFAULT_CATEGORIES
         *        REFERENCE_TYPE_DIRECT
         *        REFERENCE_TYPE_WORKSHEET
         *        REFERENCE_TYPE_NOT_USED
         *        REFERENCE_TYPE_ERROR_REPORTED
         */
        public byte ReferenceType
        {
            get
            {
                return field_2_referenceType;
            }
            set
            {
                this.field_2_referenceType = value;
            }
        }


        /**
         * Get the options field for the LinkedData record.
         */
        public short Options
        {
            get
            {
                return field_3_options;
            }
            set
            {
                this.field_3_options = value;
            }
        }

        /**
         * Get the index number fmt record field for the LinkedData record.
         */
        public short IndexNumberFmtRecord
        {
            get
            {
                return field_4_indexNumberFmtRecord;
            }
            set
            {
                this.field_4_indexNumberFmtRecord = value;
            }
        }

        /**
         * Get the formula of link field for the LinkedData record.
         */
        public Ptg[] FormulaOfLink
        {
            get
            {
                return field_5_formulaOfLink.Tokens;
            }
            set
            {
                this.field_5_formulaOfLink = Formula.Create(value);
            }
        }


        /**
         * true if this object has a custom number format
         * @return  the custom number format field value.
         */
        public bool IsCustomNumberFormat
        {
            get
            {
                return customNumberFormat.IsSet(field_3_options);
            }
            set
            {
                field_3_options = customNumberFormat.SetShortBoolean(field_3_options, value);
            }
        }
    }

}