
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
    using NPOI.Util;


    /*
     * The font index record indexes into the font table for the text record.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    //
    /// <summary>
    /// The FontX record specifies the font for a given text element. 
    /// The Font record referenced by iFont can exist in this chart sheet substream or the workbook.
    /// </summary>
    public class FontXRecord
       : StandardRecord
    {
        public const short sid = 0x1026;
        private short field_1_fontIndex;


        public FontXRecord()
        {

        }

        /**
         * Constructs a FontIndex record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public FontXRecord(RecordInputStream in1)
        {
            field_1_fontIndex = in1.ReadShort();
        }


        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[FONTX]\n");
            buffer.Append("    .fontIndex            = ")
                .Append("0x").Append(HexDump.ToHex(FontIndex))
                .Append(" (").Append(FontIndex).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/FONTX]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_fontIndex);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return  2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            FontXRecord rec = new FontXRecord();

            rec.field_1_fontIndex = field_1_fontIndex;
            return rec;
        }


        /// <summary>
        /// specifies the font to use for subsequent records.
        /// This font can either be the default font of the chart, part of the collection of Font records following 
        /// the FrtFontList record, or part of the collection of Font records in the globals substream. 
        /// If iFont is 0x0000, this record specifies the default font of the chart. 
        /// If iFont is less than or equal to the number of Font records in the globals substream, 
        ///     iFont is a one-based index to a Font record in the globals substream. 
        /// Otherwise iFont is a one-based index into the collection of Font records in this chart sheet substream 
        ///     where the index is equal to iFont ¨C n, where n is the number of Font records in the globals substream.
        /// </summary>
        public short FontIndex
        {
            get { return field_1_fontIndex; }
            set { this.field_1_fontIndex = value; }
        }



    }
}
