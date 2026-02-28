
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
    using System;
    using System.Text;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.Util;


    /*
     * Describes a linked data record.  This record referes to the series data or text.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    //
    /// <summary>
    /// The BRAI record specifies a reference to data in a sheet (1) that is used by a part of a series, 
    /// legend entry, trendline or error bars.
    /// </summary>
    public class BRAIRecord : StandardRecord
    {
        private BitField customNumberFormat = BitFieldFactory.GetInstance(0x1);

        public const short sid = 0x1051;
        private byte field_1_linkType;
        public const byte LINK_TYPE_TITLE_OR_TEXT = 0;
        public const byte LINK_TYPE_VALUES = 1;
        public const byte LINK_TYPE_CATEGORIES = 2;
        public const byte LINK_TYPE_BUBBLESIZE_VALUE = 3;
        private byte field_2_referenceType;
        public const byte REFERENCE_TYPE_DEFAULT_CATEGORIES = 0;
        public const byte REFERENCE_TYPE_DIRECT = 1;
        public const byte REFERENCE_TYPE_WORKSHEET = 2;
        public const byte REFERENCE_TYPE_NOT_USED = 3;
        public const byte REFERENCE_TYPE_ERROR_REPORTED = 4;
        private short field_3_options;
        private short field_4_indexNumberFmtRecord;

        /// <summary>
        /// A ChartParsedFormula structure that specifies the formula (section 2.2.2) that specifies the reference.
        /// </summary>
        private Formula field_5_formulaOfLink;


        public BRAIRecord()
        {

        }

        /**
         * Constructs a LinkedData record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public BRAIRecord(RecordInputStream in1)
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
            buffer.Append("    .linkType             = ")
                .Append(HexDump.ByteToHex(LinkType)).Append('\n');                
            buffer.Append(Environment.NewLine);
            buffer.Append("    .referenceType        = ").Append(HexDump.ByteToHex(ReferenceType)).Append('\n');
            buffer.Append(Environment.NewLine);
            buffer.Append("    .options              = ").Append(HexDump.ShortToHex(Options)).Append('\n');
            buffer.Append(Environment.NewLine);
            buffer.Append("         .customNumberFormat       = ").Append(IsCustomNumberFormat).Append('\n');
            buffer.Append("    .indexNumberFmtRecord = ")
                .Append(HexDump.ShortToHex(IndexNumberFmtRecord)).Append('\n');
            buffer.Append(Environment.NewLine);
            buffer.Append("    .formulaOfLink        = ");
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

        /**
         * Size of record (exluding 4 byte header)
         */
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

        public override Object Clone()
        {
            BRAIRecord rec = new BRAIRecord();

            rec.field_1_linkType = field_1_linkType;
            rec.field_2_referenceType = field_2_referenceType;
            rec.field_3_options = field_3_options;
            rec.field_4_indexNumberFmtRecord = field_4_indexNumberFmtRecord;
            rec.field_5_formulaOfLink = field_5_formulaOfLink.Copy();
            return rec;
        }




        /*
         * Get the link type field for the LinkedData record.
         *
         * @return  One of 
         *        LINK_TYPE_TITLE_OR_TEXT
         *        LINK_TYPE_VALUES
         *        LINK_TYPE_CATEGORIES
         */
        //
        /// <summary>
        /// specifies the part of the series, trendline, or error bars the referenced data specifies.
        /// </summary>
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

        /*
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
            set { this.field_2_referenceType = value; }
        }

        /*
         * Get the options field for the LinkedData record.
         */
        public short Options
        {
            get
            {
                return field_3_options;
            }
            set { this.field_3_options = value; }
        }

        /*
         * Get the index number fmt record field for the LinkedData record.
         */

        //
        /// <summary>
        /// specifies the number format to use for the data.
        /// </summary>
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


        /*
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

        /*
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