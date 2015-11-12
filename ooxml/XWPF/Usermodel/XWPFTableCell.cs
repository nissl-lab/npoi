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
    using System.Text;
    using System.Xml;
    /**
     * Represents a Cell within a {@link XWPFTable}. The
     *  Cell is the thing that holds the actual content (paragraphs etc)
     */
    public class XWPFTableCell : IBody, ICell
    {
        private CT_Tc ctTc;
        protected List<XWPFParagraph> paragraphs = null;
        protected List<XWPFTable> tables = null;
        protected List<IBodyElement> bodyElements = null;
        protected IBody part;
        private XWPFTableRow tableRow = null;

        // Create a map from this XWPF-level enum to the STVerticalJc.Enum values
        public enum XWPFVertAlign { TOP, CENTER, BOTH, BOTTOM };
        private static Dictionary<XWPFVertAlign, ST_VerticalJc> alignMap;
        // Create a map from the STVerticalJc.Enum values to the XWPF-level enums
        private static Dictionary<ST_VerticalJc, XWPFVertAlign> stVertAlignTypeMap;

        static XWPFTableCell()
        {
            // populate enum maps
            alignMap = new Dictionary<XWPFVertAlign, ST_VerticalJc>();
            alignMap.Add(XWPFVertAlign.TOP, ST_VerticalJc.top);
            alignMap.Add(XWPFVertAlign.CENTER, ST_VerticalJc.center);
            alignMap.Add(XWPFVertAlign.BOTH, ST_VerticalJc.both);
            alignMap.Add(XWPFVertAlign.BOTTOM, ST_VerticalJc.bottom);

            stVertAlignTypeMap = new Dictionary<ST_VerticalJc, XWPFVertAlign>();
            stVertAlignTypeMap.Add(ST_VerticalJc.top, XWPFVertAlign.TOP);
            stVertAlignTypeMap.Add(ST_VerticalJc.center, XWPFVertAlign.CENTER);
            stVertAlignTypeMap.Add(ST_VerticalJc.both, XWPFVertAlign.BOTH);
            stVertAlignTypeMap.Add(ST_VerticalJc.bottom, XWPFVertAlign.BOTTOM);

        }

        /**
         * If a table cell does not include at least one block-level element, then this document shall be considered corrupt
         */
        public XWPFTableCell(CT_Tc cell, XWPFTableRow tableRow, IBody part)
        {
            this.ctTc = cell;
            this.part = part;
            this.tableRow = tableRow;
            // NB: If a table cell does not include at least one block-level element, then this document shall be considered corrupt.
            if(cell.GetPList().Count<1)
                cell.AddNewP();
            bodyElements = new List<IBodyElement>();
            paragraphs = new List<XWPFParagraph>();
            tables = new List<XWPFTable>();
            foreach (object o in ctTc.Items)
            {
                if (o is CT_P)
                {
                    XWPFParagraph p = new XWPFParagraph((CT_P)o, this);
                    paragraphs.Add(p);
                    bodyElements.Add(p);
                }
                if (o is CT_Tbl)
                {
                    XWPFTable t = new XWPFTable((CT_Tbl)o, this);
                    tables.Add(t);
                    bodyElements.Add(t);
                }
                if (o is CT_SdtBlock)
                {
                    XWPFSDT c = new XWPFSDT((CT_SdtBlock)o, this);
                    bodyElements.Add(c);
                }
                if (o is CT_SdtRun)
                {
                    XWPFSDT c = new XWPFSDT((CT_SdtRun)o, this);
                    bodyElements.Add(c);
                }
            }
        }



        public CT_Tc GetCTTc()
        {
            return ctTc;
        }

        /**
         * returns an Iterator with paragraphs and tables
         * @see NPOI.XWPF.UserModel.IBody#getBodyElements()
         */
        public IList<IBodyElement> BodyElements
        {
            get
            {
                return bodyElements.AsReadOnly();
            }
        }

        public void SetParagraph(XWPFParagraph p)
        {
            if (ctTc.SizeOfPArray() == 0) {
                ctTc.AddNewP();
            }
            ctTc.SetPArray(0, p.GetCTP());
        }

        /**
         * returns a list of paragraphs
         */
        public IList<XWPFParagraph> Paragraphs
        {
            get
            {
                return paragraphs;
            }
        }

        /**
         * Add a Paragraph to this Table Cell
         * @return The paragraph which was Added
         */
        public XWPFParagraph AddParagraph()
        {
            XWPFParagraph p = new XWPFParagraph(ctTc.AddNewP(), this);
            AddParagraph(p);
            return p;
        }

        /**
         * add a Paragraph to this TableCell
         * @param p the paragaph which has to be Added
         */
        public void AddParagraph(XWPFParagraph p)
        {
            paragraphs.Add(p);
        }

        /**
         * Removes a paragraph of this tablecell
         * @param pos
         */
        public void RemoveParagraph(int pos)
        {
            paragraphs.RemoveAt(pos);
            ctTc.RemoveP(pos);
        }

        /**
         * if there is a corresponding {@link XWPFParagraph} of the parameter ctTable in the paragraphList of this table
         * the method will return this paragraph
         * if there is no corresponding {@link XWPFParagraph} the method will return null 
         * @param p is instance of CTP and is searching for an XWPFParagraph
         * @return null if there is no XWPFParagraph with an corresponding CTPparagraph in the paragraphList of this table
         * 		   XWPFParagraph with the correspondig CTP p
         */
        public XWPFParagraph GetParagraph(CT_P p)
        {
            foreach (XWPFParagraph paragraph in paragraphs) {
                if(p.Equals(paragraph.GetCTP())){
                    return paragraph;
                }
            }
            return null;
        }

        public void SetBorderBottom(XWPFTable.XWPFBorderType type, int size, int space, String rgbColor)
        {
            CT_TcPr ctTcPr = null;
            if (!GetCTTc().IsSetTcPr())
            {
                ctTcPr = GetCTTc().AddNewTcPr();
            }
            CT_TcBorders borders = ctTcPr.AddNewTcBorders();
            borders.bottom = new CT_Border();
            CT_Border b = borders.bottom;
            b.val = XWPFTable.xwpfBorderTypeMap[type];
            b.sz = (ulong)size;
            b.space = (ulong)space;
            b.color = (rgbColor);
        }

        public void SetText(String text)
        {
            CT_P ctP = (ctTc.SizeOfPArray() == 0) ? ctTc.AddNewP() : ctTc.GetPArray(0);
            XWPFParagraph par = new XWPFParagraph(ctP, this);
            par.CreateRun().AppendText(text);
        }

        public XWPFTableRow GetTableRow()
        {
            return tableRow;
        }

        /**
     * Set cell color. This sets some associated values; for finer control
     * you may want to access these elements individually.
     * @param rgbStr - the desired cell color, in the hex form "RRGGBB".
     */
        public void SetColor(String rgbStr)
        {
            CT_TcPr tcpr = ctTc.IsSetTcPr() ? ctTc.tcPr : ctTc.AddNewTcPr();
            CT_Shd ctshd = tcpr.IsSetShd() ? tcpr.shd : tcpr.AddNewShd();
            ctshd.color = ("auto");
            ctshd.val = (ST_Shd.clear);
            ctshd.fill = (rgbStr);
        }

        /**
         * Get cell color. Note that this method only returns the "fill" value.
         * @return RGB string of cell color
         */
        public String GetColor()
        {
            String color = null;
            CT_TcPr tcpr = ctTc.tcPr;
            if (tcpr != null)
            {
                CT_Shd ctshd = tcpr.shd;
                if (ctshd != null)
                {
                    color = ctshd.fill;
                }
            }
            return color;
        }

        /**
         * Set the vertical alignment of the cell.
         * @param vAlign - the desired alignment enum value
         */
        public void SetVerticalAlignment(XWPFVertAlign vAlign)
        {
            CT_TcPr tcpr = ctTc.IsSetTcPr() ? ctTc.tcPr : ctTc.AddNewTcPr();
            CT_VerticalJc va = tcpr.AddNewVAlign();
            va.val = (alignMap[(vAlign)]);
        }

        /**
         * Get the vertical alignment of the cell.
         * @return the cell alignment enum value
         */
        public XWPFVertAlign GetVerticalAlignment()
        {
            XWPFVertAlign vAlign = XWPFVertAlign.TOP;
            CT_TcPr tcpr = ctTc.tcPr;
            if (ctTc != null)
            {
                CT_VerticalJc va = tcpr.vAlign;
                vAlign = stVertAlignTypeMap[(va.val)];
            }
            return vAlign;
        }

        /**
         * add a new paragraph at position of the cursor
         * @param cursor
         * @return the inserted paragraph
         */
        public XWPFParagraph InsertNewParagraph(/*XmlCursor*/ XmlDocument cursor)
        {
            /*if(!isCursorInTableCell(cursor))
                return null;
            
            String uri = CTP.type.Name.NamespaceURI;
            String localPart = "p";
            cursor.BeginElement(localPart,uri);
            cursor.ToParent();
            CTP p = (CTP)cursor.Object;
            XWPFParagraph newP = new XWPFParagraph(p, this);
            XmlObject o = null;
            while(!(o is CTP)&&(cursor.ToPrevSibling())){
                o = cursor.Object;
            }
            if((!(o is CTP)) || (CTP)o == p){
                paragraphs.Add(0, newP);
            }
            else{
                int pos = paragraphs.IndexOf(getParagraph((CTP)o))+1;
                paragraphs.Add(pos,newP);
            }
            int i=0;
            cursor.ToCursor(p.NewCursor());
            while(cursor.ToPrevSibling()){
                o =cursor.Object;
                if(o is CTP || o is CTTbl)
                    i++;
            }
            bodyElements.Add(i, newP);
            cursor.ToCursor(p.NewCursor());
            cursor.ToEndToken();
            return newP;*/
            throw new NotImplementedException();
        }

        public XWPFTable InsertNewTbl(/*XmlCursor*/ XmlDocument cursor)
        {
            /*if(isCursorInTableCell(cursor)){
                String uri = CTTbl.type.Name.NamespaceURI;
                String localPart = "tbl";
                cursor.BeginElement(localPart,uri);
                cursor.ToParent();
                CTTbl t = (CTTbl)cursor.Object;
                XWPFTable newT = new XWPFTable(t, this);
                cursor.RemoveXmlContents();
                XmlObject o = null;
                while(!(o is CTTbl)&&(cursor.ToPrevSibling())){
                    o = cursor.Object;
                }
                if(!(o is CTTbl)){
                    tables.Add(0, newT);
                }
                else{
                    int pos = tables.IndexOf(getTable((CTTbl)o))+1;
                    tables.Add(pos,newT);
                }
                int i=0;
                cursor = t.NewCursor();
                while(cursor.ToPrevSibling()){
                    o =cursor.Object;
                    if(o is CTP || o is CTTbl)
                        i++;
                }
                bodyElements.Add(i, newT);
                cursor = t.NewCursor();
                cursor.ToEndToken();
                return newT;
            }
            return null;*/
            throw new NotImplementedException();
        }

        /**
         * verifies that cursor is on the right position
         */
        private bool IsCursorInTableCell(/*XmlCursor*/XmlDocument cursor)
        {
            /*XmlCursor verify = cursor.NewCursor();
            verify.ToParent();
            if(verify.Object == this.ctTc){
                return true;
            }
            return false;*/
            throw new NotImplementedException();
        }



        /**
         * @see NPOI.XWPF.UserModel.IBody#getParagraphArray(int)
         */
        public XWPFParagraph GetParagraphArray(int pos)
        {
            if (pos > 0 && pos < paragraphs.Count)
            {
                return paragraphs[(pos)];
            }
            return null;
        }

        /**
         * Get the to which the TableCell belongs
         * 
         * @see NPOI.XWPF.UserModel.IBody#getPart()
         */
        public POIXMLDocumentPart Part
        {
            get
            {
                return tableRow.GetTable().Part;
            }
        }


        /** 
         * @see NPOI.XWPF.UserModel.IBody#getPartType()
         */
        public BodyType PartType
        {
            get
            {
                return BodyType.TABLECELL;
            }
        }


        /**
         * Get a table by its CTTbl-Object
         * @see NPOI.XWPF.UserModel.IBody#getTable(org.Openxmlformats.schemas.wordProcessingml.x2006.main.CTTbl)
         */
        public XWPFTable GetTable(CT_Tbl ctTable)
        {
            for(int i=0; i<tables.Count; i++){
                if(this.Tables[(i)].GetCTTbl() == ctTable) return Tables[(i)]; 
            }
            return null;
        }


        /** 
         * @see NPOI.XWPF.UserModel.IBody#getTableArray(int)
         */
        public XWPFTable GetTableArray(int pos)
        {
            if (pos >=0 && pos < tables.Count)
            {
                return tables[pos];
            }
            return null;
        }


        /** 
         * @see NPOI.XWPF.UserModel.IBody#getTables()
         */
        public IList<XWPFTable> Tables
        {
            get
            {
                return tables.AsReadOnly();
            }
        }


        /**
         * inserts an existing XWPFTable to the arrays bodyElements and tables
         * @see NPOI.XWPF.UserModel.IBody#insertTable(int, NPOI.XWPF.UserModel.XWPFTable)
         */
        public void InsertTable(int pos, XWPFTable table)
        {
            bodyElements.Insert(pos, table);
            int i;
            for (i = 0; i < ctTc.GetTblList().Count; i++) {
                CT_Tbl tbl = ctTc.GetTblArray(i);
                if(tbl == table.GetCTTbl()){
                    break;
                }
            }
            tables.Insert(i, table);
        }

        public String GetText()
        {
            StringBuilder text = new StringBuilder();
            foreach (XWPFParagraph p in paragraphs)
            {
                text.Append(p.Text);
            }
            return text.ToString();
        }

        /**
     * extracts all text recursively through embedded tables and embedded SDTs
     */
        public String GetTextRecursively()
        {

            StringBuilder text = new StringBuilder();
            for (int i = 0; i < bodyElements.Count; i++)
            {
                bool isLast = (i == bodyElements.Count - 1) ? true : false;
                AppendBodyElementText(text, bodyElements[i], isLast);
            }

            return text.ToString();
        }

        private void AppendBodyElementText(StringBuilder text, IBodyElement e, bool isLast)
        {
            if (e is XWPFParagraph)
            {
                text.Append(((XWPFParagraph)e).Text);
                if (isLast == false)
                {
                    text.Append('\t');
                }
            }
            else if (e is XWPFTable)
            {
                XWPFTable eTable = (XWPFTable)e;
                foreach (XWPFTableRow row in eTable.Rows)
                {
                    foreach (XWPFTableCell cell in row.GetTableCells())
                    {
                        IList<IBodyElement> localBodyElements = cell.BodyElements;
                        for (int i = 0; i < localBodyElements.Count; i++)
                        {
                            bool localIsLast = (i == localBodyElements.Count - 1) ? true : false;
                            AppendBodyElementText(text, localBodyElements[i], localIsLast);
                        }
                    }
                }

                if (isLast == false)
                {
                    text.Append('\n');
                }
            }
            else if (e is XWPFSDT)
            {
                text.Append(((XWPFSDT)e).Content.Text);
                if (isLast == false)
                {
                    text.Append('\t');
                }
            }
        }
        /**
         * Get the TableCell which belongs to the TableCell
         */
        public XWPFTableCell GetTableCell(CT_Tc cell)
        {
            if (!(cell.Parent is CT_Row))
                return null;
            CT_Row row = (CT_Row)cell.Parent;

            if (!(row.Parent is CT_Tbl))
            {
                return null;
            }
            CT_Tbl tbl = (CT_Tbl)row.Parent;
            XWPFTable table = GetTable(tbl);
            if (table == null)
            {
                return null;
            }
            XWPFTableRow tableRow = table.GetRow(row);
            if (tableRow == null)
            {
                return null;
            }
            return tableRow.GetTableCell(cell);
        }

        public XWPFDocument GetXWPFDocument()
        {
            return part.GetXWPFDocument();
        }
    }// end class
}
