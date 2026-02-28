/* ====================================================================
   Copyright 2002-2004   Apache Software Foundation

   Licensed Under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.Util
{

    using NPOI.Util;
    using System.Collections;
    using NPOI.HSSF.Record;
    /**
     * <p>Title: HSSFCellRangeAddress</p>
     * <p>Description:
     *          Implementation of the cell range Address lists,like Is described in
     *          OpenOffice.org's Excel Documentation .
     *          In BIFF8 there Is a common way to store absolute cell range Address
     *          lists in several records (not formulas). A cell range Address list
     *          consists of a field with the number of ranges and the list of the range
     *          Addresses. Each cell range Address (called an AddR structure) Contains
     *          4 16-bit-values.</p>
     * <p>Copyright: Copyright (c) 2004</p>
     * <p>Company: </p>
     * @author Dragos Buleandra (dragos.buleandra@trade2b.ro)
     * @version 2.0-pre
     */

    public class HSSFCellRangeAddress
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(HSSFCellRangeAddress));

        /**
         * Number of following AddR structures
         */
        private short field_Addr_number;

        /**
         * List of AddR structures. Each structure represents a cell range
         */
        private ArrayList field_regions_list;

        public HSSFCellRangeAddress()
        {

        }

        /**
         * Construct a new HSSFCellRangeAddress object and Sets its fields appropriately .
         * Even this Isn't an Excel record , I kept the same behavior for reading/writing
         * the object's data as for a regular record .
         * 
         * @param in the RecordInputstream to read the record from
         */
        public HSSFCellRangeAddress(RecordInputStream in1)
        {
            this.FillFields(in1);
        }

        public void FillFields(RecordInputStream in1)
        {
            this.field_Addr_number = in1.ReadShort();
            this.field_regions_list = new ArrayList(this.field_Addr_number);

            for (int k = 0; k < this.field_Addr_number; k++)
            {
                short first_row = in1.ReadShort();
                short first_col = in1.ReadShort();

                short last_row = first_row;
                short last_col = first_col;
                if (in1.Remaining >= 4)
                {
                    last_row = in1.ReadShort();
                    last_col = in1.ReadShort();
                }
                else
                {
                    // Ran out of data
                    // For now, Issue a warning, finish, and 
                    //  hope for the best....
                    logger.Log(POILogger.WARN, "Ran out of data reading cell references for DVRecord");
                    k = this.field_Addr_number;
                }

                AddrStructure region = new AddrStructure(first_row, first_col, last_row, last_col);
                this.field_regions_list.Add(region);
            }
        }

        /**
         * Get the number of following AddR structures.
         * The number of this structures Is automatically Set when reading an Excel file
         * and/or increased when you manually Add a new AddR structure .
         * This Is the reason there Isn't a Set method for this field .
         * @return number of AddR structures
         */
        public short AddRStructureNumber
        {
            get { return this.field_Addr_number; }
        }

        /**
         * Add an AddR structure .
         * @param first_row - the upper left hand corner's row
         * @param first_col - the upper left hand corner's col
         * @param last_row  - the lower right hand corner's row
         * @param last_col  - the lower right hand corner's col
         * @return the index of this AddR structure
         */
        public int AddAddRStructure(short first_row, short first_col, short last_row, short last_col)
        {
            if (this.field_regions_list == null)
            {
                //just to be sure :-)
                this.field_Addr_number = 0;
                this.field_regions_list = new ArrayList(10);
            }
            AddrStructure region = new AddrStructure(first_row, last_row, first_col, last_col);

            this.field_regions_list.Add(region);
            this.field_Addr_number++;
            return this.field_Addr_number;
        }

        /**
         * Remove the AddR structure stored at the passed in index
         * @param index The AddR structure's index
         */
        public void RemoveAddRStructureAt(int index)
        {
            this.field_regions_list.Remove(index);
            this.field_Addr_number--;
        }

        /**
         * return the AddR structure at the given index.
         * @return AddrStructure representing
         */
        public AddrStructure GetAddRStructureAt(int index)
        {
            return (AddrStructure)this.field_regions_list[index];
        }

        public int Serialize(int offSet, byte[] data)
        {
            int pos = 2;

            LittleEndian.PutShort(data, offSet, this.AddRStructureNumber);
            for (int k = 0; k < this.AddRStructureNumber; k++)
            {
                AddrStructure region = this.GetAddRStructureAt(k);
                LittleEndian.PutShort(data, offSet + pos, region.FirstRow);
                pos += 2;
                LittleEndian.PutShort(data, offSet + pos, region.LastRow);
                pos += 2;
                LittleEndian.PutShort(data, offSet + pos, region.FirstColumn);
                pos += 2;
                LittleEndian.PutShort(data, offSet + pos, region.LastColumn);
                pos += 2;
            }
            return this.Size;
        }

        public int Size
        {
            get { return 2 + this.field_Addr_number * 8; }
        }

        public class AddrStructure
        {
            private short _first_row;
            private short _first_col;
            private short _last_row;
            private short _last_col;

            public AddrStructure(short first_row, short last_row, short first_col, short last_col)
            {
                this._first_row = first_row;
                this._last_row = last_row;
                this._first_col = first_col;
                this._last_col = last_col;
            }

            /**
             * Get the upper left hand corner column number
             * @return column number for the upper left hand corner
             */
            public short FirstColumn
            {
                get{return this._first_col;}
                set { this._first_col = value; }
            }

            /**
             * Get the upper left hand corner row number
             * @return row number for the upper left hand corner
             */
            public short FirstRow
            {
                get{return this._first_row;}
                set { this._first_row = value; }
            }

            /**
             * Get the lower right hand corner column number
             * @return column number for the lower right hand corner
             */
            public short LastColumn
            {
                get{return this._last_col;}
                set { this._last_col = value; }
            }

            /**
             * Get the lower right hand corner row number
             * @return row number for the lower right hand corner
             */
            public short LastRow
            {
                get { return this._last_row; }
                set { this._last_row = value; }
            }


        }
    }

}