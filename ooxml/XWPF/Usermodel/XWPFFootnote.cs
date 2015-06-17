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
    using System.Collections.Generic;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.Xml;

    public class XWPFFootnote : IEnumerator<XWPFParagraph>, IBody
    {
        private List<XWPFParagraph> paragraphs = new List<XWPFParagraph>();
        private List<XWPFTable> tables = new List<XWPFTable>();
        private List<XWPFPictureData> pictures = new List<XWPFPictureData>();
        private List<IBodyElement> bodyElements = new List<IBodyElement>();

        private CT_FtnEdn ctFtnEdn;
        private XWPFFootnotes footnotes;
        private XWPFDocument document;

        public XWPFFootnote(CT_FtnEdn note, XWPFFootnotes xFootnotes)
        {
            footnotes = xFootnotes;
            ctFtnEdn = note;
            document = xFootnotes.GetXWPFDocument();
            Init();
        }

        public XWPFFootnote(XWPFDocument document, CT_FtnEdn body)
        {
            ctFtnEdn = body;
            this.document = document;
            Init();
        }
        private void Init()
        {
            //copied from XWPFDocument...should centralize this code
            //to avoid duplication       
            foreach (object o in ctFtnEdn.Items)
            {

                if (o is CT_P)
                {
                    XWPFParagraph p = new XWPFParagraph((CT_P)o, this);
                    bodyElements.Add(p);
                    paragraphs.Add(p);
                }
                else if (o is CT_Tbl)
                {
                    XWPFTable t = new XWPFTable((CT_Tbl)o, this);
                    bodyElements.Add(t);
                    tables.Add(t);
                }
                else if (o is CT_SdtBlock)
                {
                    XWPFSDT c = new XWPFSDT((CT_SdtBlock)o, this);
                    bodyElements.Add(c);
                }
            }
        }
        public IList<XWPFParagraph> Paragraphs
        {
            get
            {
                return paragraphs;
            }
        }
        public IEnumerator<XWPFParagraph> GetEnumerator()
        {
            return paragraphs.GetEnumerator();
        }

        public IList<XWPFTable> Tables
        {
            get
            {
                return tables;
            }
        }

        public IList<XWPFPictureData> Pictures
        {
            get
            {
                return pictures;
            }
        }

        public IList<IBodyElement> BodyElements
        {
            get
            {
                return bodyElements;
            }
        }

        public CT_FtnEdn GetCTFtnEdn()
        {
            return ctFtnEdn;
        }

        public void SetCTFtnEdn(CT_FtnEdn footnote)
        {
            ctFtnEdn = footnote;
        }

         /// <summary>
         /// 
         /// </summary>
         /// <param name="pos">position in table array</param>
        /// <returns>The table at position pos</returns>
        public XWPFTable GetTableArray(int pos)
        {
            if (pos > 0 && pos < tables.Count)
            {
                return tables[(pos)];
            }
            return null;
        }

        /// <summary>
        /// inserts an existing XWPFTable to the arrays bodyElements and tables
        /// </summary>
        /// <param name="pos"></param>
        /// <param name="table"></param>
        public void InsertTable(int pos, XWPFTable table)
        {
            bodyElements.Insert(pos, table);
            int i;
            for (i = 0; i < ctFtnEdn.GetTblList().Count; i++) {
                CT_Tbl tbl = ctFtnEdn.GetTblArray(i);
                if(tbl == table.GetCTTbl()){
                    break;
                }
            }
            tables.Insert(i, table);
        }

        /**
         * if there is a corresponding {@link XWPFTable} of the parameter ctTable in the tableList of this header
         * the method will return this table
         * if there is no corresponding {@link XWPFTable} the method will return null 
         * @param ctTable
         * @see NPOI.XWPF.UserModel.IBody#getTable(CTTbl ctTable)
         */
        public XWPFTable GetTable(CT_Tbl ctTable)
        {
            foreach (XWPFTable table in tables) {
                if(table==null)
                    return null;
                if(table.GetCTTbl().Equals(ctTable))
                    return table;	
            }
            return null;
        }

        /**
         * if there is a corresponding {@link XWPFParagraph} of the parameter ctTable in the paragraphList of this header or footer
         * the method will return this paragraph
         * if there is no corresponding {@link XWPFParagraph} the method will return null 
         * @param p is instance of CTP and is searching for an XWPFParagraph
         * @return null if there is no XWPFParagraph with an corresponding CTPparagraph in the paragraphList of this header or footer
         * 		   XWPFParagraph with the correspondig CTP p
         * @see NPOI.XWPF.UserModel.IBody#getParagraph(CTP p)
         */
        public XWPFParagraph GetParagraph(CT_P p)
        {
            foreach (XWPFParagraph paragraph in paragraphs) {
                if(paragraph.GetCTP().Equals(p))
                    return paragraph;
            }
            return null;
        }
        /// <summary>
        /// Returns the paragraph that holds the text of the header or footer.
        /// </summary>
        /// <param name="pos"></param>
        /// <returns></returns>
        public XWPFParagraph GetParagraphArray(int pos)
        {

            return paragraphs[pos];
        }

        /// <summary>
        /// Get the TableCell which belongs to the TableCell
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public XWPFTableCell GetTableCell(CT_Tc cell)
        {
            object obj = cell.Parent;
            if (!(obj is CT_Row))
                return null;

            CT_Row row = (CT_Row)obj;
            if (!(row.Parent is CT_Tbl))
                return null;

            CT_Tbl tbl = (CT_Tbl)row.Parent;
            XWPFTable table = GetTable(tbl);
            if(table == null){
                return null;
            }
            XWPFTableRow tableRow = table.GetRow(row);
            if(row == null){
                return null;
            }
            return tableRow.GetTableCell(cell);
        }

        /**
         * verifies that cursor is on the right position
         * @param cursor
         */
        private bool IsCursorInFtn(/*XmlCursor*/XmlDocument cursor)
        {
            /*XmlCursor verify = cursor.NewCursor();
            verify.ToParent();
            if(verify.Object == this.ctFtnEdn){
                return true;
            }
            return false;*/
            throw new NotImplementedException();
        }

        public POIXMLDocumentPart Owner
        {
            get
            {
                return footnotes;
            }
        }

        /**
         * 
         * @param cursor
         * @return the inserted table
         * @see NPOI.XWPF.UserModel.IBody#insertNewTbl(XmlCursor cursor)
         */
        public XWPFTable InsertNewTbl(/*XmlCursor*/XmlDocument cursor)
        {
            /*if(isCursorInFtn(cursor)){
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
         * add a new paragraph at position of the cursor
         * @param cursor
         * @return the inserted paragraph
         * @see NPOI.XWPF.UserModel.IBody#insertNewParagraph(XmlCursor cursor)
         */
        public XWPFParagraph InsertNewParagraph(/*XmlCursor*/XmlDocument cursor)
        {
            /*if(isCursorInFtn(cursor)){
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
                return newP;
            }
            return null;*/
            throw new NotImplementedException();
        }

        /**
         * add a new table to the end of the footnote
         * @param table
         * @return the Added XWPFTable
         */
        public XWPFTable AddNewTbl(CT_Tbl table)
        {
            CT_Tbl newTable = ctFtnEdn.AddNewTbl();
            newTable.Set(table);
            XWPFTable xTable = new XWPFTable(newTable, this);
            tables.Add(xTable);
            return xTable;
        }

        /**
         * add a new paragraph to the end of the footnote
         * @param paragraph
         * @return the Added XWPFParagraph
         */
        public XWPFParagraph AddNewParagraph(CT_P paragraph)
        {
            CT_P newPara = ctFtnEdn.AddNewP(paragraph);
            //newPara.Set(paragraph);
            XWPFParagraph xPara = new XWPFParagraph(newPara, this);
            paragraphs.Add(xPara);
            return xPara;
        }

        /**
         * @see NPOI.XWPF.UserModel.IBody#getXWPFDocument()
         */
        public XWPFDocument GetXWPFDocument()
        {
            return document;
        }

        /**
         * returns the Part, to which the body belongs, which you need for Adding relationship to other parts
         * @see NPOI.XWPF.UserModel.IBody#getPart()
         */
        public POIXMLDocumentPart Part
        {
            get
            {
                return footnotes;
            }
        }

        /**
         * Get the PartType of the body
         * @see NPOI.XWPF.UserModel.IBody#getPartType()
         */
        public BodyType PartType
        {
            get
            {
                return BodyType.FOOTNOTE;
            }
        }

        public XWPFParagraph Current
        {
            get { throw new NotImplementedException(); }
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        object System.Collections.IEnumerator.Current
        {
            get { throw new NotImplementedException(); }
        }

        public bool MoveNext()
        {
            throw new NotImplementedException();
        }

        public void Reset()
        {
            throw new NotImplementedException();
        }
    }
}
