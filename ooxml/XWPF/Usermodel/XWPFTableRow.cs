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
     * @author gisellabronzetti
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


        public CT_Row GetCtRow()
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

        public XWPFTableCell GetCell(int pos)
        {
            if (pos >= 0 && pos < ctRow.SizeOfTcArray())
            {
                return GetTableCells()[(pos)];
            }
            return null;
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
         * @param height
         */
        public void SetHeight(int height)
        {
            CT_TrPr properties = GetTrPr();
            CT_Height h = properties.SizeOfTrHeightArray() == 0 ? properties.AddNewTrHeight() : properties.GetTrHeightArray(0);
            h.val = (ulong)height;
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
        public int GetHeight()
        {
            CT_TrPr properties = GetTrPr();
            return properties.SizeOfTrHeightArray() == 0 ? 0 : (int)properties.GetTrHeightArray(0).val;
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
        public List<XWPFTableCell> GetTableCells()
        {
            if (tableCells == null)
            {
                List<XWPFTableCell> cells = new List<XWPFTableCell>();
                foreach (CT_Tc tableCell in ctRow.GetTcList())
                {
                    cells.Add(new XWPFTableCell(tableCell, this, table.Body));
                }
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

    }// end class

}