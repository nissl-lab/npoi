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
    using NPOI.OpenXml4Net.OPC;
    using System.IO;
    using NPOI.Util;
    using System.Text;
    using NPOI.OpenXmlFormats.Wordprocessing;

    /**
     * Parent of XWPF headers and footers
     */
    public abstract class XWPFHeaderFooter : POIXMLDocumentPart, IBody
    {
        protected List<XWPFParagraph> paragraphs = new List<XWPFParagraph>(1);
        protected List<XWPFTable> tables = new List<XWPFTable>(1);
        protected List<XWPFPictureData> pictures = new List<XWPFPictureData>();
        protected List<IBodyElement> bodyElements = new List<IBodyElement>(1);

        protected CT_HdrFtr headerFooter;
        protected XWPFDocument document;

        public XWPFHeaderFooter(XWPFDocument doc, CT_HdrFtr hdrFtr)
        {
            if (doc == null)
            {
                throw new NullReferenceException();
            }

            document = doc;
            headerFooter = hdrFtr;
            ReadHdrFtr();
        }

        protected XWPFHeaderFooter()
        {

            //headerFooter = new CT_HdrFtr();
            //ReadHdrFtr();
        }

        public XWPFHeaderFooter(POIXMLDocumentPart parent, PackagePart part, PackageRelationship rel) :
            base(parent, part, rel)
        {
            ;
            this.document = (XWPFDocument)GetParent();

            if (this.document == null)
            {
                throw new NullReferenceException();
            }
        }


        internal override void OnDocumentRead()
        {
            foreach (POIXMLDocumentPart poixmlDocumentPart in GetRelations())
            {
                if (poixmlDocumentPart is XWPFPictureData)
                {
                    XWPFPictureData xwpfPicData = (XWPFPictureData)poixmlDocumentPart;
                    pictures.Add(xwpfPicData);
                    document.RegisterPackagePictureData(xwpfPicData);
                }
            }
        }


        public CT_HdrFtr _getHdrFtr()
        {
            return headerFooter;
        }

        public IList<IBodyElement> BodyElements
        {
            get
            {
                return bodyElements.AsReadOnly();
            }
        }

        /**
         * Returns the paragraph(s) that holds
         *  the text of the header or footer.
         * Normally there is only the one paragraph, but
         *  there could be more in certain cases, or 
         *  a table.
         */
        public IList<XWPFParagraph> Paragraphs
        {
            get
            {
                return paragraphs.AsReadOnly();
            }
        }


        /**
         * Return the table(s) that holds the text
         *  of the header or footer, for complex cases
         *  where a paragraph isn't used.
         * Normally there's just one paragraph, but some
         *  complex headers/footers have a table or two
         *  in Addition. 
         */
        public IList<XWPFTable> Tables
        {
            get
            {
                return tables.AsReadOnly();
            }
        }



        /**
         * Returns the textual content of the header/footer,
         *  by flattening out the text of its paragraph(s)
         */
        public String Text
        {
            get
            {
                StringBuilder t = new StringBuilder();

                for (int i = 0; i < paragraphs.Count; i++)
                {
                    if (!paragraphs[(i)].IsEmpty)
                    {
                        String text = paragraphs[i].Text;
                        if (text != null && text.Length > 0)
                        {
                            t.Append(text);
                            t.Append('\n');
                        }
                    }
                }

                IList<XWPFTable> tables = this.Tables;
                for (int i = 0; i < tables.Count; i++)
                {
                    String text = tables[(i)].Text;
                    if (text != null && text.Length > 0)
                    {
                        t.Append(text);
                        t.Append('\n');
                    }
                }
                foreach (IBodyElement bodyElement in BodyElements)
                {
                    if (bodyElement is XWPFSDT)
                    {
                        t.Append(((XWPFSDT)bodyElement).Content.Text + '\n');
                    }
                }
                return t.ToString();
            }
        }

        /**
         * Set a new headerFooter
         */
        public void SetHeaderFooter(CT_HdrFtr headerFooter)
        {
            this.headerFooter = headerFooter;
            ReadHdrFtr();
        }

        /**
         * if there is a corresponding {@link XWPFTable} of the parameter ctTable in the tableList of this header
         * the method will return this table
         * if there is no corresponding {@link XWPFTable} the method will return null 
         * @param ctTable
         */
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

        /**
         * if there is a corresponding {@link XWPFParagraph} of the parameter ctTable in the paragraphList of this header or footer
         * the method will return this paragraph
         * if there is no corresponding {@link XWPFParagraph} the method will return null 
         * @param p is instance of CTP and is searching for an XWPFParagraph
         * @return null if there is no XWPFParagraph with an corresponding CTPparagraph in the paragraphList of this header or footer
         * 		   XWPFParagraph with the correspondig CTP p
         */
        public XWPFParagraph GetParagraph(CT_P p)
        {
            foreach (XWPFParagraph paragraph in paragraphs)
            {
                if (paragraph.GetCTP().Equals(p))
                    return paragraph;
            }
            return null;
        }

        /**
         * Returns the paragraph that holds
         *  the text of the header or footer.
         */
        public XWPFParagraph GetParagraphArray(int pos)
        {

            return paragraphs[(pos)];
        }

        /**
         * Get a List of all Paragraphs
         * @return a list of {@link XWPFParagraph} 
         */
        public List<XWPFParagraph> GetListParagraph()
        {
            return paragraphs;
        }

        public IList<XWPFPictureData> AllPictures
        {
            get
            {
                return pictures.AsReadOnly();
            }
        }

        /**
         * Get all Pictures in this package
         * @return all Pictures in this package
         */
        public IList<XWPFPictureData> AllPackagePictures
        {
            get
            {
                return document.AllPackagePictures;
            }
        }

        /**
         * Adds a picture to the document.
         *
         * @param pictureData       The picture data
         * @param format            The format of the picture.
         *
         * @return the index to this picture (0 based), the Added picture can be obtained from {@link #getAllPictures()} .
         * @throws InvalidFormatException 
         */
        public String AddPictureData(byte[] pictureData, int format)
        {
            XWPFPictureData xwpfPicData = document.FindPackagePictureData(pictureData, format);
            POIXMLRelation relDesc = XWPFPictureData.RELATIONS[format];

            if (xwpfPicData == null)
            {
                /* Part doesn't exist, create a new one */
                int idx = document.GetNextPicNameNumber(format);
                xwpfPicData = (XWPFPictureData)CreateRelationship(relDesc, XWPFFactory.GetInstance(), idx);
                /* write bytes to new part */
                PackagePart picDataPart = xwpfPicData.GetPackagePart();
                Stream out1 = null;
                try
                {
                    out1 = picDataPart.GetOutputStream();
                    out1.Write(pictureData, 0, pictureData.Length);
                }
                catch (IOException e)
                {
                    throw new POIXMLException(e);
                }
                finally
                {
                    try
                    {
                        if (out1 != null)
                            out1.Close();
                    }
                    catch (IOException)
                    {
                        // ignore
                    }
                }

                document.RegisterPackagePictureData(xwpfPicData);
                pictures.Add(xwpfPicData);
                return GetRelationId(xwpfPicData);
            }
            else if (!GetRelations().Contains(xwpfPicData))
            {
                /*
                 * Part already existed, but was not related so far. Create
                 * relationship to the already existing part and update
                 * POIXMLDocumentPart data.
                 */
                PackagePart picDataPart = xwpfPicData.GetPackagePart();
                // TODO add support for TargetMode.EXTERNAL relations.
                TargetMode targetMode = TargetMode.Internal;
                PackagePartName partName = picDataPart.PartName;
                String relation = relDesc.Relation;
                PackageRelationship relShip = GetPackagePart().AddRelationship(partName, targetMode, relation);
                String id = relShip.Id;
                AddRelation(id, xwpfPicData);
                pictures.Add(xwpfPicData);
                return id;
            }
            else
            {
                /* Part already existed, Get relation id and return it */
                return GetRelationId(xwpfPicData);
            }
        }

        /**
         * Adds a picture to the document.
         *
         * @param is                The stream to read image from
         * @param format            The format of the picture.
         *
         * @return the index to this picture (0 based), the Added picture can be obtained from {@link #getAllPictures()} .
         * @throws InvalidFormatException 
         * @ 
         */
        public String AddPictureData(Stream is1, int format)
        {
            byte[] data = IOUtils.ToByteArray(is1);
            return AddPictureData(data, format);
        }

        /**
         * returns the PictureData by blipID
         * @param blipID
         * @return XWPFPictureData of a specificID
         * @throws Exception 
         */
        public XWPFPictureData GetPictureDataByID(String blipID)
        {
            POIXMLDocumentPart relatedPart = GetRelationById(blipID);
            if (relatedPart != null && relatedPart is XWPFPictureData)
            {
                return (XWPFPictureData)relatedPart;
            }
            return null;
        }

        /**
         * add a new paragraph at position of the cursor
         * @param cursor
         * @return the inserted paragraph
         */
        /*public XWPFParagraph insertNewParagraph(XmlCursor cursor)
        {
            if (isCursorInHdrF(cursor))
            {
                String uri = CTP.type.Name.NamespaceURI;
                String localPart = "p";
                cursor.BeginElement(localPart, uri);
                cursor.ToParent();
                CTP p = (CTP)cursor.Object;
                XWPFParagraph newP = new XWPFParagraph(p, this);
                XmlObject o = null;
                while (!(o is CTP) && (cursor.ToPrevSibling()))
                {
                    o = cursor.Object;
                }
                if ((!(o is CTP)) || (CTP)o == p)
                {
                    paragraphs.Add(0, newP);
                }
                else
                {
                    int pos = paragraphs.IndexOf(getParagraph((CTP)o)) + 1;
                    paragraphs.Add(pos, newP);
                }
                int i = 0;
                cursor.ToCursor(p.NewCursor());
                while (cursor.ToPrevSibling())
                {
                    o = cursor.Object;
                    if (o is CTP || o is CTTbl)
                        i++;
                }
                bodyElements.Add(i, newP);
                cursor.ToCursor(p.NewCursor());
                cursor.ToEndToken();
                return newP;
            }
            return null;
        }*/


        /**
         * 
         * @param cursor
         * @return the inserted table
         */
        /*public XWPFTable insertNewTbl(XmlCursor cursor)
        {
            if (isCursorInHdrF(cursor))
            {
                String uri = CTTbl.type.Name.NamespaceURI;
                String localPart = "tbl";
                cursor.BeginElement(localPart, uri);
                cursor.ToParent();
                CTTbl t = (CTTbl)cursor.Object;
                XWPFTable newT = new XWPFTable(t, this);
                cursor.RemoveXmlContents();
                XmlObject o = null;
                while (!(o is CTTbl) && (cursor.ToPrevSibling()))
                {
                    o = cursor.Object;
                }
                if (!(o is CTTbl))
                {
                    tables.Add(0, newT);
                }
                else
                {
                    int pos = tables.IndexOf(getTable((CTTbl)o)) + 1;
                    tables.Add(pos, newT);
                }
                int i = 0;
                cursor = t.NewCursor();
                while (cursor.ToPrevSibling())
                {
                    o = cursor.Object;
                    if (o is CTP || o is CTTbl)
                        i++;
                }
                bodyElements.Add(i, newT);
                cursor = t.NewCursor();
                cursor.ToEndToken();
                return newT;
            }
            return null;
        }*/

        /**
         * verifies that cursor is on the right position
         * @param cursor
         */
        /*private bool IsCursorInHdrF(XmlCursor cursor)
        {
            XmlCursor verify = cursor.NewCursor();
            verify.ToParent();
            if (verify.Object == this.headerFooter)
            {
                return true;
            }
            return false;
        }*/


        public POIXMLDocumentPart Owner
        {
            get
            {
                return this;
            }
        }

        /**
         * Returns the table at position pos
         * @see NPOI.XWPF.UserModel.IBody#getTableArray(int)
         */
        public XWPFTable GetTableArray(int pos)
        {

            if (pos > 0 && pos < tables.Count)
            {
                return tables[(pos)];
            }
            return null;
        }

        /**
         * inserts an existing XWPFTable to the arrays bodyElements and tables
         * @param pos
         * @param table
         */
        public void InsertTable(int pos, XWPFTable table)
        {
            bodyElements.Insert(pos, table);
            int i;
            for (i = 0; i < headerFooter.GetTblList().Count; i++)
            {
                CT_Tbl tbl = headerFooter.GetTblArray(i);
                if (tbl == table.GetCTTbl())
                {
                    break;
                }
            }
            tables.Insert(i, table);
        }

        public void ReadHdrFtr()
        {
            bodyElements = new List<IBodyElement>();
            paragraphs = new List<XWPFParagraph>();
            tables = new List<XWPFTable>();
            // parse the document with cursor and add
            // the XmlObject to its lists
            foreach (object o in headerFooter.Items)
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
            }
            /*XmlCursor cursor = headerFooter.NewCursor();
            cursor.SelectPath("./*");
            while (cursor.ToNextSelection())
            {
                XmlObject o = cursor.Object;
                if (o is CTP)
                {
                    XWPFParagraph p = new XWPFParagraph((CTP)o, this);
                    paragraphs.Add(p);
                    bodyElements.Add(p);
                }
                if (o is CTTbl)
                {
                    XWPFTable t = new XWPFTable((CTTbl)o, this);
                    tables.Add(t);
                    bodyElements.Add(t);
                }
            }
            cursor.Dispose();*/
        }

        /**
         * Get the TableCell which belongs to the TableCell
         * @param cell
         */
        public XWPFTableCell GetTableCell(CT_Tc cell)
        {
            /*XmlCursor cursor = cell.NewCursor();
            cursor.ToParent();
            XmlObject o = cursor.Object;
            if(!(o is CTRow)){
                return null;
            }
            CTRow row = (CTRow)o;
            cursor.ToParent();
            o = cursor.Object;
            cursor.Dispose();
            if(! (o is CTTbl)){
                return null;
            }
            CTTbl tbl = (CTTbl) o;
            XWPFTable table = GetTable(tbl);
            if(table == null){
                return null;
            }
            XWPFTableRow tableRow = table.GetRow(row);
            if(row == null){
                return null;
            }
            return tableRow.GetTableCell(cell);*/
            throw new NotImplementedException();
        }

        public void SetXWPFDocument(XWPFDocument doc)
        {
            document = doc;
        }
        public XWPFDocument GetXWPFDocument()
        {
            if (document != null)
            {
                return document;
            }
            else
            {
                return (XWPFDocument)GetParent();
            }
        }

        /**
         * returns the Part, to which the body belongs, which you need for Adding relationship to other parts
         * @see NPOI.XWPF.UserModel.IBody#getPart()
         */
        public POIXMLDocumentPart Part
        {
            get
            {
                return this;
            }
        }

        #region IBody ≥…‘±


        public virtual BodyType PartType
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public XWPFParagraph InsertNewParagraph(System.Xml.XmlDocument cursor)
        {
            throw new NotImplementedException();
        }

        public XWPFTable InsertNewTbl(System.Xml.XmlDocument cursor)
        {
            throw new NotImplementedException();
        }

        #endregion
    }//end class

}