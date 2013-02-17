
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


    /**
     * The series list record defines the series Displayed as an overlay to the main chart record.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class SeriesListRecord
       : StandardRecord
    {
        public const short sid = 0x1016;
        private short[] field_1_seriesNumbers;


        public SeriesListRecord(short[] seriesNumbers)
        {
            field_1_seriesNumbers = seriesNumbers;
        }

        /**
         * Constructs a SeriesList record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public SeriesListRecord(RecordInputStream in1)
        {
            int nItems = in1.ReadUShort();
    	    short[] ss = new short[nItems];
    	    for (int i = 0; i < nItems; i++) {
			    ss[i] = in1.ReadShort();
    			
		    }
            field_1_seriesNumbers = ss;

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[SERIESLIST]\n");
            buffer.Append("    .seriesNumbers        = ")
                .Append(" (").Append(SeriesNumbers).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/SERIESLIST]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            int nItems = field_1_seriesNumbers.Length;
            out1.WriteShort(nItems);
            for (int i = 0; i < nItems; i++)
            {
                out1.WriteShort(field_1_seriesNumbers[i]);
            }
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return field_1_seriesNumbers.Length * 2 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            return new SeriesListRecord((short[])field_1_seriesNumbers.Clone());
        }




        /**
         * Get the series numbers field for the SeriesList record.
         */
        public short[] SeriesNumbers
        {
            get { return field_1_seriesNumbers; }
            set { this.field_1_seriesNumbers = value; }
        }


    }
}