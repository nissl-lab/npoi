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
namespace NPOI.XWPF.UserModel
{
    using System;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.Collections.Generic;


    /**
     * A row within an {@link XWPFTable}. Rows mostly just have
     *  sizings and stylings, the interesting content lives inside
     *  the child {@link XWPFTableCell}s
     */
    public class XWPFTableRow
    {

        private CT_Row ctRow;
        private XWPFTable table;
        private List<XWPFTableCell> tableCells;

        public XWPFTableRow(CT_Row row, XWPFTable table)
        {
            this.table = table;
            this.ctRow = row;
            GetTableCells();
        }


        public CT_Row GetCTRow()
        {
            return ctRow;
        }

        /**
         * create a new XWPFTableCell and add it to the tableCell-list of this tableRow
         * @return the newly Created XWPFTableCell
         */
        public XWPFTableCell CreateCell()
        {
            XWPFTableCell tableCell = new XWPFTableCell(ctRow.AddNewTc(), this, table.Body);
            tableCells.Add(tableCell);
            return tableCell;
        }
        public void MergeCells(int startIndex, int endIndex)
        {
            if (startIndex >= endIndex)
            {
                throw new ArgumentOutOfRangeException("Start index must be smaller than end index");
            }
            if (startIndex < 0 || endIndex >= this.tableCells.Count)
            {
                throw new ArgumentOutOfRangeException("Invalid start index and end index");
            }
            XWPFTableCell startCell = this.GetCell(startIndex);
            //remove merged cells
            for (int i = endIndex; i >startIndex; i--)
                this.RemoveCell(i);
            
            if (!startCell.GetCTTc().IsSetTcPr())
            {
                startCell.GetCTTc().AddNewTcPr();
            }
            CT_TcPr tcPr = startCell.GetCTTc().tcPr;
            if(tcPr.gridSpan==null)
                tcPr.AddNewGridspan();
            CT_DecimalNumber gridspan = tcPr.gridSpan;
            gridspan.val = (endIndex - startIndex+1).ToString();
        }
        public XWPFTableCell GetCell(int pos)
        {
            if (pos >= 0 && pos < ctRow.SizeOfTcArray())
            {
                return GetTableCells()[(pos)];
            }
            return null;
        }
        public void RemoveCell(int pos)
        {
            if (pos >= 0 && pos < ctRow.SizeOfTcArray())
            {
                tableCells.RemoveAt(pos);
                ctRow.RemoveTc(pos);
            }
        }
        /**
         * Adds a new TableCell at the end of this tableRow
         */
        public XWPFTableCell AddNewTableCell()
        {
            CT_Tc cell = ctRow.AddNewTc();
            XWPFTableCell tableCell = new XWPFTableCell(cell, this, table.Body);
            tableCells.Add(tableCell);
            return tableCell;
        }


        /**
         * This element specifies the height of the current table row within the
         * current table. This height shall be used to determine the resulting
         * height of the table row, which may be absolute or relative (depending on
         * its attribute values). If omitted, then the table row shall automatically
         * resize its height to the height required by its contents (the equivalent
         * of an hRule value of auto).
         *
         * @return height
         */
        public int Height
        {
            get
            {
                CT_TrPr properties = GetTrPr();
                return properties.SizeOfTrHeightArray() == 0 ? 0 : (int)properties.GetTrHeightArray(0).val;
            }
            set
            {
                CT_TrPr properties = GetTrPr();
                CT_Height h = properties.SizeOfTrHeightArray() == 0 ? properties.AddNewTrHeight() : properties.GetTrHeightArray(0);
                h.val = (ulong)value;
            }
        }


        private CT_TrPr GetTrPr()
        {
            return (ctRow.IsSetTrPr()) ? ctRow.trPr : ctRow.AddNewTrPr();
        }

        public XWPFTable GetTable()
        {
            return table;
        }

        /**
     * create and return a list of all XWPFTableCell
     * who belongs to this row
     * @return a list of {@link XWPFTableCell} 
     */
        public List<ICell> GetTableICells()
        {

            List<ICell> cells = new List<ICell>();
            //Can't use ctRow.getTcList because that only gets table cells
            //Can't use ctRow.getSdtList because that only gets sdts that are at cell level

            foreach(object o in ctRow.Items)
            {
                if (o is CT_Tc)
                {
                    cells.Add(new XWPFTableCell((CT_Tc)o, this, table.Body));
                }
                else if (o is CT_SdtCell)
                {
                    cells.Add(new XWPFSDTCell((CT_SdtCell)o, this, table.Body));
                }
            }
            return cells;
        }

        /**
         * create and return a list of all XWPFTableCell
         * who belongs to this row
         * @return a list of {@link XWPFTableCell} 
         */
        public List<XWPFTableCell> GetTableCells()
        {
            if (tableCells == null)
            {
                List<XWPFTableCell> cells = new List<XWPFTableCell>();
                foreach (CT_Tc tableCell in ctRow.GetTcList())
                {
                    cells.Add(new XWPFTableCell(tableCell, this, table.Body));
                }

                //TODO: it is possible to have an SDT that contains a cell in within a row
                //need to modify this code so that it pulls out SDT wrappers around cells, too.


                this.tableCells = cells;
            }
            return tableCells;
        }

        /**
         * returns the XWPFTableCell which belongs to the CTTC cell
         * if there is no XWPFTableCell which belongs to the parameter CTTc cell null will be returned
         */
        public XWPFTableCell GetTableCell(CT_Tc cell)
        {
            for (int i = 0; i < tableCells.Count; i++)
            {
                if (tableCells[(i)].GetCTTc() == cell) 
                    return tableCells[(i)];
            }
            return null;
        }

        /**
         * Return true if the "can't split row" value is true. The logic for this
         * attribute is a little unusual: a TRUE value means DON'T allow rows to
         * split, FALSE means allow rows to split.
         * @return true if rows can't be split, false otherwise.
         */
        public bool IsCantSplitRow
        {
            get
            {
                bool isCant = false;
                CT_TrPr trpr = GetTrPr();
                if (trpr.SizeOfCantSplitArray() > 0)
                {
                    CT_OnOff onoff = trpr.GetCantSplitList()[0];
                    isCant = onoff.val;
                }
                return isCant;
            }
            set 
            {
                CT_TrPr trpr = GetTrPr();
                CT_OnOff onoff = trpr.AddNewCantSplit();
                onoff.val = value;
            }
        }

        /**
         * Return true if a table's header row should be repeated at the top of a
         * table split across pages.
         * @return true if table's header row should be repeated at the top of each
         *         page of table, false otherwise.
         */
        public bool IsRepeatHeader
        {
            get
            {
                bool repeat = false;
                CT_TrPr trpr = GetTrPr();
                if (trpr.SizeOfTblHeaderArray() > 0)
                {
                    CT_OnOff rpt = trpr.GetTblHeaderList()[0];
                    repeat = rpt.val;
                }
                return repeat;
            }
            set 
            {
                CT_TrPr trpr = GetTrPr();
                CT_OnOff onoff = trpr.AddNewTblHeader();
                onoff.val = value;
            }
        }
    }// end class

}