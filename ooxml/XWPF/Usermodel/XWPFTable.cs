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
    using System.Text;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.Collections.Generic;

    /**
     * Sketch of XWPFTable class. Only table's text is being hold.
     * <p/>
     * Specifies the contents of a table present in the document. A table is a Set
     * of paragraphs (and other block-level content) arranged in rows and columns.
     *
     * @author Yury Batrakov (batrakov at gmail.com)
     */
    public class XWPFTable : IBodyElement
    {

        protected StringBuilder text = new StringBuilder();
        private CT_Tbl ctTbl;
        protected List<XWPFTableRow> tableRows;
        protected List<String> styleIDs;
        protected IBody part;

        public XWPFTable(CT_Tbl table, IBody part, int row, int col)
            : this(table, part)
        {

            for (int i = 0; i < row; i++)
            {
                XWPFTableRow tabRow = (GetRow(i) == null) ? CreateRow() : GetRow(i);
                for (int k = 0; k < col; k++)
                {
                    XWPFTableCell tabCell = (tabRow.GetCell(k) == null) ? tabRow
                            .CreateCell() : null;
                }
            }
        }

        public XWPFTable(CT_Tbl table, IBody part)
        {
            this.part = part;
            this.ctTbl = table;

            tableRows = new List<XWPFTableRow>();
            /*
            // is an empty table: I add one row and one column as default
            if (table.SizeOfTrArray() == 0)
                CreateEmptyTable(table);

            foreach (CTRow row in table.TrList) {
                StringBuilder rowText = new StringBuilder();
                XWPFTableRow tabRow = new XWPFTableRow(row, this);
                tableRows.Add(tabRow);
                foreach (CTTc cell in row.TcList) {
                    foreach (CTP ctp in cell.PList) {
                        XWPFParagraph p = new XWPFParagraph(ctp, part);
                        if (rowText.Length() > 0) {
                            rowText.Append('\t');
                        }
                        rowText.Append(p.Text);
                    }
                }
                if (rowText.Length() > 0) {
                    this.text.Append(rowText);
                    this.text.Append('\n');
                }
            }*/
            throw new NotImplementedException();
        }

        private void CreateEmptyTable(CT_Tbl table)
        {
            // MINIMUM ELEMENTS FOR A TABLE
            /*table.AddNewTr().AddNewTc().addNewP();

            CTTblPr tblpro = table.AddNewTblPr();
            tblpro.AddNewTblW().W=(new Bigint("0"));
            tblpro.TblW.Type=(STTblWidth.AUTO);

            // layout
            // tblpro.AddNewTblLayout().Type=(STTblLayoutType.AUTOFIT);

            // borders
            CTTblBorders borders = tblpro.AddNewTblBorders();
            borders.AddNewBottom().Val=(STBorder.SINGLE);
            borders.AddNewInsideH().Val=(STBorder.SINGLE);
            borders.AddNewInsideV().Val=(STBorder.SINGLE);
            borders.AddNewLeft().Val=(STBorder.SINGLE);
            borders.AddNewRight().Val=(STBorder.SINGLE);
            borders.AddNewTop().Val=(STBorder.SINGLE);
            */
            /*
           * CTTblGrid tblgrid=table.AddNewTblGrid();
           * tblgrid.AddNewGridCol().W=(new Bigint("2000"));
           */
            throw new NotImplementedException();
            GetRows();
        }

        /**
         * @return ctTbl object
         */

        public CT_Tbl GetCTTbl()
        {
            return ctTbl;
        }

        /**
         * @return text
         */
        public String GetText()
        {
            return text.ToString();
        }

        public void AddNewRowBetween(int start, int end)
        {
            // TODO
        }

        /**
         * add a new column for each row in this table
         */
        public void AddNewCol()
        {
            /*if (ctTbl.SizeOfTrArray() == 0) {
                CreateRow();
            }
            for (int i = 0; i < ctTbl.SizeOfTrArray(); i++) {
                XWPFTableRow tabRow = new XWPFTableRow(ctTbl.GetTrArray(i), this);
                tabRow.CreateCell();
            }*/
            throw new NotImplementedException();
        }

        /**
         * create a new XWPFTableRow object with as many cells as the number of columns defined in that moment
         *
         * @return tableRow
         */
        public XWPFTableRow CreateRow()
        {
            /*int sizeCol = ctTbl.SizeOfTrArray() > 0 ? ctTbl.GetTrArray(0)
                    .sizeOfTcArray() : 0;
            XWPFTableRow tabRow = new XWPFTableRow(ctTbl.AddNewTr(), this);
            AddColumn(tabRow, sizeCol);
            tableRows.Add(tabRow);
            return tabRow;*/
            throw new NotImplementedException();
        }

        /**
         * @param pos - index of the row
         * @return the row at the position specified or null if no rows is defined or if the position is greather than the max size of rows array
         */
        public XWPFTableRow GetRow(int pos)
        {
            /*if (pos >= 0 && pos < ctTbl.SizeOfTrArray()) {
                //return new XWPFTableRow(ctTbl.GetTrArray(pos));
                return GetRows().Get(pos);
            }
            return null;*/
            throw new NotImplementedException();
        }


        /**
         * @param width
         */
        public void SetWidth(int width)
        {
            /*CTTblPr tblPr = GetTrPr();
            CTTblWidth tblWidth = tblPr.IsSetTblW() ? tblPr.TblW : tblPr
                    .AddNewTblW();
            tblWidth.W=(new Bigint("" + width));
             * */
            throw new NotImplementedException();
        }

        /**
         * @return width value
         */
        public int GetWidth()
        {
            /*CTTblPr tblPr = GetTrPr();
            return tblPr.IsSetTblW() ? tblPr.TblW.W.IntValue() : -1;
             */
            throw new NotImplementedException();
        }

        /**
         * @return number of rows in table
         */
        public int GetNumberOfRows()
        {
            //return ctTbl.SizeOfTrArray();
            throw new NotImplementedException();
        }

        private CT_TblPr GetTrPr()
        {
            /*return (ctTbl.TblPr != null) ? ctTbl.TblPr : ctTbl
                    .AddNewTblPr();
             */
            throw new NotImplementedException();
        }

        private void AddColumn(XWPFTableRow tabRow, int sizeCol)
        {
            if (sizeCol > 0)
            {
                for (int i = 0; i < sizeCol; i++)
                {
                    tabRow.CreateCell();
                }
            }
        }

        /**
         * Get the StyleID of the table
         * @return	style-ID of the table
         */
        public String GetStyleID()
        {
            //return ctTbl.TblPr.TblStyle.Val;
            throw new NotImplementedException();
        }

        /**
         * add a new Row to the table
         * 
         * @param row	the row which should be Added
         */
        public void AddRow(XWPFTableRow row)
        {
            /*ctTbl.AddNewTr();
            ctTbl.TrArray=(getNumberOfRows()-1, row.CtRow);
            tableRows.Add(row);*/
            throw new NotImplementedException();
        }

        /**
         * add a new Row to the table
         * at position pos
         * @param row	the row which should be Added
         */
        public bool AddRow(XWPFTableRow row, int pos){
        //if(pos >= 0 && pos <= tableRows.Size()){
        //    ctTbl.InsertNewTr(pos);
        //    ctTbl.TrArray=(pos,row.CtRow);
        //    tableRows.Add(pos, row);
        //    return true;
        //}
            throw new NotImplementedException();
    	//return false;
    }

        /**
         * inserts a new tablerow 
         * @param pos
         * @return  the inserted row
         */
        public XWPFTableRow insertNewTableRow(int pos)
        {
            /*if(pos >= 0 && pos <= tableRows.Size()){
                CTRow row = ctTbl.InsertNewTr(pos);
                XWPFTableRow tableRow = new XWPFTableRow(row, this);
                tableRows.Add(pos, tableRow);
                return tableRow;
            }
            return null;*/
            throw new NotImplementedException();
        }


        /**
         * Remove a row at position pos from the table
         * @param pos	position the Row in the Table
         */
        public bool RemoveRow(int pos)
        {
            /*if (pos >= 0 && pos < tableRows.Size()) {
                ctTbl.RemoveTr(pos);
                tableRows.Remove(pos);
                return true;
            }
            return false;*/
            throw new NotImplementedException();
        }

        public List<XWPFTableRow> GetRows()
        {
            return tableRows;
        }


        /**
         * returns the type of the BodyElement Table
         * @see NPOI.XWPF.UserModel.IBodyElement#getElementType()
         */
        public BodyElementType GetElementType()
        {
            return BodyElementType.TABLE;
        }

        public IBody GetBody()
        {
            return part;
        }

        /**
         * returns the part of the bodyElement
         * @see NPOI.XWPF.UserModel.IBody#getPart()
         */
        public POIXMLDocumentPart GetPart()
        {
            if (part != null)
            {
                return part.GetPart();
            }
            return null;
        }

        /**
         * returns the partType of the bodyPart which owns the bodyElement
         * @see NPOI.XWPF.UserModel.IBody#getPartType()
         */
        public BodyType GetPartType()
        {
            return part.GetPartType();
        }

        /**
         * returns the XWPFRow which belongs to the CTRow row
         * if this row is not existing in the table null will be returned
         */
        public XWPFTableRow GetRow(CT_Row row)
        {
            /*for(int i=0; i<getRows().Count; i++){
                if(getRows().Get(i).CtRow== row) return GetRow(i); 
            }
            return null;*/
            throw new NotImplementedException();
        }
    }// end class

}