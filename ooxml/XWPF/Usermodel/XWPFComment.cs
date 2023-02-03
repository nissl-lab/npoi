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
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Xml;

    /**
     * Sketch of XWPF comment class
     */
    public class XWPFComment : IBody
    {
        protected CT_Comment ctComment;
        protected XWPFComments comments;
        protected XWPFDocument document;
        private List<XWPFParagraph> paragraphs = new List<XWPFParagraph>();
        private List<XWPFTable> tables = new List<XWPFTable>();
        private List<IBodyElement> bodyElements = new List<IBodyElement>();

        public XWPFComment(CT_Comment ctComment, XWPFComments comments)
        {
            this.comments = comments;
            this.ctComment = ctComment;
            this.document = comments.GetXWPFDocument();
            Init();
        }

        protected void Init()
        {
            foreach (var o in ctComment.Items)
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

        public String Id
        {
            get
            {
                return GetId();
            }
        }

        public String Author
        {
            get
            {
                return GetAuthor();
            }
            set
            {
                SetAuthor(value);
            }
        }

        public String Text
        {
            get
            {
                return GetText();
            }
        }

        public String Initials
        {
            get
            {
                return GetInitials();
            }
            set
            {
                SetInitials(value);
            }
        }

        public string Date
        {
            get
            {
                return GetDate();
            }
            set
            {
                SetDate(value);
            }
        }

        /**
         * Get the Part to which the comment belongs, which you need for adding
         * relationships to other parts
         *
         * @return {@link POIXMLDocumentPart} that contains the comment.
         * @see org.apache.poi.xwpf.usermodel.IBody#getPart()
         */
        public POIXMLDocumentPart Part
        {
            get { return comments; }
        }

        /**
         * Get the part type {@link BodyType} of the comment.
         *
         * @return The {@link BodyType} value.
         * @see org.apache.poi.xwpf.usermodel.IBody#getPartType()
         */
        public BodyType PartType
        {
            get { return BodyType.COMMENT; }
        }

        /**
         * Gets the body elements ({@link IBodyElement}) of the comment.
         *
         * @return List of body elements.
         */
        public IList<IBodyElement> BodyElements
        {
            get { return bodyElements.AsReadOnly(); }
        }

        /**
         * Returns the paragraph(s) that holds the text of the comment.
         */
        public IList<XWPFParagraph> Paragraphs
        {
            get { return paragraphs.AsReadOnly(); }
        }

        /**
         * Get the list of {@link XWPFTable}s in the comment.
         *
         * @return List of tables
         */
        public IList<XWPFTable> Tables
        {
            get { return tables.AsReadOnly(); }
        }

        public XWPFParagraph GetParagraph(CT_P p)
        {
            foreach (XWPFParagraph paragraph in paragraphs)
            {
                if (paragraph.GetCTP().Equals(p))
                    return paragraph;
            }
            return null;
        }

        public XWPFTable GetTable(CT_Tbl ctTable)
        {
            foreach (XWPFTable table in tables)
            {
                if (table == null)
                    return null;
                if (table.GetCTTbl().Equals(ctTable))
                    return table;
            }
            return null;
        }

        public XWPFParagraph GetParagraphArray(int pos)
        {
            if (pos >= 0 && pos < paragraphs.Count)
            {
                return paragraphs[pos];
            }
            return null;
        }

        public XWPFTable GetTableArray(int pos)
        {
            if (pos >= 0 && pos < tables.Count)
            {
                return tables[pos];
            }
            return null;
        }

        public XWPFParagraph InsertNewParagraph(XmlDocument cursor)
        {
            //if (isCursorInCmt(cursor))
            //{
            //    String uri = CT_P.type.getName().getNamespaceURI();
            //    String localPart = "p";
            //    cursor.beginElement(localPart, uri);
            //    cursor.toParent();
            //    CT_P p = (CT_P)cursor.getObject();
            //    XWPFParagraph newP = new XWPFParagraph(p, this);
            //    XmlObject o = null;
            //    while (!(o is CT_P) && (cursor.toPrevSibling()))
            //    {
            //        o = cursor.getObject();
            //    }
            //    if ((!(o is CT_P)) || o == p)
            //    {
            //        paragraphs.add(0, newP);
            //    }
            //    else
            //    {
            //        int pos = paragraphs.indexOf(getParagraph((CT_P)o)) + 1;
            //        paragraphs.add(pos, newP);
            //    }
            //    int i = 0;
            //    using (XmlCursor p2 = p.newCursor())
            //    {
            //        cursor.toCursor(p2);
            //    }
            //    while (cursor.toPrevSibling())
            //    {
            //        o = cursor.getObject();
            //        if (o is CT_P || o is CT_Tbl)
            //            i++;
            //    }
            //    bodyElements.add(i, newP);
            //    using (XmlCursor p2 = p.newCursor())
            //    {
            //        cursor.toCursor(p2);
            //        cursor.toEndToken();
            //    }
            //    return newP;
            //}
            //return null;
            throw new NotImplementedException();
        }

        //private bool isCursorInCmt(XmlCursor cursor)
        //{
        //    using (XmlCursor verify = cursor.newCursor())
        //    {
        //        verify.toParent();
        //        return (verify.getObject() == this.ctComment);
        //    }
        //}

        public XWPFTable InsertNewTbl(XmlDocument cursor)
        {
            //if (isCursorInCmt(cursor))
            //{
            //    String uri = CT_Tbl.type.getName().getNamespaceURI();
            //    String localPart = "tbl";
            //    cursor.beginElement(localPart, uri);
            //    cursor.toParent();
            //    CT_Tbl t = (CT_Tbl)cursor.getObject();
            //    XWPFTable newT = new XWPFTable(t, this);
            //    cursor.removeXmlContents();
            //    XmlObject o = null;
            //    while (!(o is CT_Tbl) && (cursor.toPrevSibling()))
            //    {
            //        o = cursor.getObject();
            //    }
            //    if (!(o is CT_Tbl))
            //    {
            //        tables.add(0, newT);
            //    }
            //    else
            //    {
            //        int pos = tables.indexOf(getTable((CT_Tbl)o)) + 1;
            //        tables.add(pos, newT);
            //    }
            //    int i = 0;
            //    using (XmlCursor cursor2 = t.newCursor())
            //    {
            //        while (cursor2.toPrevSibling())
            //        {
            //            o = cursor2.getObject();
            //            if (o is CT_P || o is CT_Tbl)
            //            {
            //                i++;
            //            }
            //        }
            //    }
            //    bodyElements.add(i, newT);
            //    using (XmlCursor cursor2 = t.newCursor())
            //    {
            //        cursor.toCursor(cursor2);
            //        cursor.toEndToken();
            //    }
            //    return newT;
            //}
            //return null;
            throw new NotImplementedException();
        }

        public void InsertTable(int pos, XWPFTable table)
        {
            bodyElements.Insert(pos, table);
            int i = 0;
            foreach (CT_Tbl tbl in ctComment.GetTblList())
            {
                if (tbl == table.GetCTTbl())
                {
                    break;
                }
                i++;
            }
            tables.Insert(i, table);
        }

        public XWPFTableCell GetTableCell(CT_Tc cell)
        {
            //XmlObject o;
            //CT_Row row;
            //using (final XmlCursor cursor = cell.newCursor()) {
            //    cursor.toParent();
            //    o = cursor.getObject();
            //    if (!(o is CT_Row))
            //    {
            //        return null;
            //    }
            //    row = (CT_Row)o;
            //    cursor.toParent();
            //    o = cursor.getObject();
            //}
            //if (!(o is CT_Tbl))
            //{
            //    return null;
            //}
            //CT_Tbl tbl = (CT_Tbl)o;
            //XWPFTable table = getTable(tbl);
            //if (table == null)
            //{
            //    return null;
            //}
            //XWPFTableRow tableRow = table.getRow(row);
            //return tableRow.getTableCell(cell);
            throw new NotImplementedException();
        }

        /**
         * Get the {@link XWPFDocument} the comment is part of.
         *
         * @see org.apache.poi.xwpf.usermodel.IBody#getXWPFDocument()
         */
        public XWPFDocument GetXWPFDocument()
        {
            return document;
        }

        public String GetText()
        {
            StringBuilder text = new StringBuilder();
            foreach (XWPFParagraph p in paragraphs)
            {
                if (text.Length > 0)
                {
                    text.Append("\n");
                }
                text.Append(p.Text);
            }
            return text.ToString();
        }

        public XWPFParagraph CreateParagraph()
        {
            XWPFParagraph paragraph = new XWPFParagraph(ctComment.AddNewP(), this);
            paragraphs.Add(paragraph);
            bodyElements.Add(paragraph);
            return paragraph;
        }

        //public void RemoveParagraph(XWPFParagraph paragraph)
        //{
        //    if (paragraphs.Contains(paragraph))
        //    {
        //        CT_P ctP = paragraph.GetCTP();
        //        using (XmlCursor c = ctP.newCursor())
        //        {
        //            c.removeXml();
        //        }
        //        paragraphs.Remove(paragraph);
        //        bodyElements.Remove(paragraph);
        //    }
        //}

        //public void RemoveTable(XWPFTable table)
        //{
        //    if (tables.Contains(table))
        //    {
        //        CT_Tbl ctTbl = table.GetCTTbl();
        //        using (XmlCursor c = ctTbl.newCursor())
        //        {
        //            c.removeXml();
        //        }
        //        tables.Remove(table);
        //        bodyElements.Remove(table);
        //    }
        //}

        public XWPFTable CreateTable(int rows, int cols)
        {
            XWPFTable table = new XWPFTable(ctComment.AddNewTbl(), this, rows, cols);
            tables.Add(table);
            bodyElements.Add(table);
            return table;
        }

        /**
         * Gets the underlying CT_Comment object for the comment.
         *
         * @return CT_Comment object
         */
        public CT_Comment GetCtComment()
        {
            return ctComment;
        }

        /**
         * The owning object for this comment
         *
         * @return The {@link XWPFComments} object that contains this comment.
         */
        public XWPFComments GetComments()
        {
            return comments;
        }

        /**
         * Get a unique identifier for the current comment. The restrictions on the
         * id attribute, if any, are defined by the parent XML element. If this
         * attribute is omitted, then the document is non-conformant.
         *
         * @return string id
         */
        public String GetId()
        {
            return ctComment.id.ToString();
        }

        /**
         * Get the author of the current comment
         *
         * @return author of the current comment
         */
        public String GetAuthor()
        {
            return ctComment.author;
        }

        /**
         * Specifies the author for the current comment If this attribute is
         * omitted, then no author shall be associated with the parent annotation
         * type.
         *
         * @param author author of the current comment
         */
        public void SetAuthor(String author)
        {
            ctComment.author = author;
        }

        /**
         * Get the initials of the author of the current comment
         *
         * @return initials the initials of the author of the current comment
         */
        public String GetInitials()
        {
            return ctComment.initials;
        }

        /**
         * Specifies the initials of the author of the current comment
         *
         * @param initials the initials of the author of the current comment
         */
        public void SetInitials(String initials)
        {
            ctComment.initials = initials;
        }

        /**
         * Get the date information of the current comment
         *
         * @return the date information for the current comment.
         */
        public string GetDate()
        {
            return ctComment.date;
        }

        /**
         * Specifies the date information for the current comment. If this attribute
         * is omitted, then no date information shall be associated with the parent
         * annotation type.
         *
         * @param date the date information for the current comment.
         */
        public void SetDate(string date)
        {
            ctComment.date = date;
        }
    }

}
