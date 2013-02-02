
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
        

/*
 * MulRKRecord.java
 *
 * Created on November 9, 2001, 4:53 PM
 */
namespace NPOI.HSSF.Record
{
    using NPOI.Util;
    using System;
    using System.Text;
    using NPOI.HSSF.Util;

    /**
     * Used to store multiple RK numbers on a row.  1 MulRk = Multiple Cell values.
     * HSSF just Converts this into multiple NUMBER records.  Read-ONLY SUPPORT!
     * REFERENCE:  PG 330 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class MulRKRecord : StandardRecord
    {
        public const short sid = 0xbd;
        private int field_1_row;
        private short field_2_first_col;
        private RkRec[] field_3_rks;
        private short field_4_last_col;

        /** Creates new MulRKRecord */

        public MulRKRecord()
        {
        }

        /**
         * Constructs a MulRK record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public MulRKRecord(RecordInputStream in1)
        {
            field_1_row = in1.ReadUShort();
            field_2_first_col = in1.ReadShort();
            field_3_rks = RkRec.ParseRKs(in1);
            field_4_last_col = in1.ReadShort();
        }

        //public short Row
        public int Row
        {
            get { return field_1_row; }
        }

        /**
         * starting column (first cell this holds in the row)
         * @return first column number
         */

        public short FirstColumn
        {
            get { return field_2_first_col; }
        }

        /**
         * ending column (last cell this holds in the row)
         * @return first column number
         */

        public short LastColumn
        {
            get { return field_4_last_col; }
        }

        /**
         * Get the number of columns this Contains (last-first +1)
         * @return number of columns (last - first +1)
         */

        public int NumColumns
        {
            get { return field_4_last_col - field_2_first_col + 1; }
        }

        /**
         * returns the xf index for column (coffset = column - field_2_first_col)
         * @return the XF index for the column
         */

        public short GetXFAt(int coffset)
        {
            return field_3_rks[coffset].xf;
        }

        /**
         * returns the rk number for column (coffset = column - field_2_first_col)
         * @return the value (decoded into a double)
         */

        public double GetRKNumberAt(int coffset)
        {
            return RKUtil.DecodeNumber(field_3_rks[coffset].rk);
        }


        //private ArrayList ParseRKs(RecordInputStream in1)
        //{
        //    ArrayList retval = new ArrayList();
        //    while ((in1.Remaining - 2) > 0)
        //    {
        //        RkRec rec = new RkRec();

        //        rec.xf = in1.ReadShort();
        //        rec.rk = in1.ReadInt();
        //        retval.Add(rec);
        //    }
        //    return retval;
        //}

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[MULRK]\n");
            buffer.Append("	.row	 = ").Append(HexDump.ShortToHex(Row)).Append("\n");
            buffer.Append(" .firstcol= ").Append(StringUtil.ToHexString(FirstColumn)).Append("\n");
            buffer.Append(" .lastcol = ").Append(StringUtil.ToHexString(LastColumn)).Append("\n");
            for (int k = 0; k < NumColumns; k++)
            {
                buffer.Append(" xf[").Append(k).Append("] = ").Append(StringUtil.ToHexString(GetXFAt(k))).Append("\n");
                buffer.Append(" rk[").Append(k).Append("] = ").Append(GetRKNumberAt(k)).Append("\n");
            }
            buffer.Append("[/MULRK]\n");
            return buffer.ToString();
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            throw new RecordFormatException("Sorry, you can't serialize MulRK in this release");
        }
        protected override int DataSize
        {
            get
            {
                throw new RecordFormatException("Sorry, you can't serialize MulRK in this release");
            }
        }
        private class RkRec
        {
            public const int ENCODED_SIZE = 6;
            public short xf;
            public int rk;

            private RkRec(RecordInputStream in1)
            {
                xf = in1.ReadShort();
                rk = in1.ReadInt();
            }

            public static RkRec[] ParseRKs(RecordInputStream in1)
            {
                int nItems = (in1.Remaining - 2) / ENCODED_SIZE;
                RkRec[] retval = new RkRec[nItems];
                for (int i = 0; i < nItems; i++)
                {
                    retval[i] = new RkRec(in1);
                }
                return retval;
            }
        }
    }

    
}