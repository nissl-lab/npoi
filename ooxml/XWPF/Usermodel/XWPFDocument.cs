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
    using System.IO;
    using NPOI.Util;
    using System.Collections.Generic;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.Xml;
    using NPOI.XWPF.Model;

    /**
     * Experimental class to do low level Processing
     *  of docx files.
     *
     * If you're using these low level classes, then you
     *  will almost certainly need to refer to the OOXML
     *  specifications from
     *  http://www.ecma-international.org/publications/standards/Ecma-376.htm
     *
     * WARNING - APIs expected to change rapidly
     */
    public class XWPFDocument : POIXMLDocument, Document, IBody
    {

        //private CTDocument1 ctDocument;
        private XWPFSettings Settings;
        /**
         * Keeps track on all id-values used in this document and included parts, like headers, footers, etc.
         */
        private IdentifierManager DrawingIdManager = new IdentifierManager(1L, 4294967295L);
        protected List<XWPFFooter> footers = new List<XWPFFooter>();
        protected List<XWPFHeader> headers = new List<XWPFHeader>();
        protected List<XWPFComment> comments = new List<XWPFComment>();
        protected List<XWPFHyperlink> hyperlinks = new List<XWPFHyperlink>();
        protected List<XWPFParagraph> paragraphs = new List<XWPFParagraph>();
        protected List<XWPFTable> tables = new List<XWPFTable>();
        protected List<IBodyElement> bodyElements = new List<IBodyElement>();
        protected List<XWPFPictureData> pictures = new List<XWPFPictureData>();
        protected Dictionary<long, List<XWPFPictureData>> packagePictures = new Dictionary<long, List<XWPFPictureData>>();
        protected Dictionary<int, XWPFFootnote> endnotes = new Dictionary<int, XWPFFootnote>();
        protected XWPFNumbering numbering;
        protected XWPFStyles styles;
        protected XWPFFootnotes footnotes;

        /** Handles the joy of different headers/footers for different pages */
        private XWPFHeaderFooterPolicy headerFooterPolicy;

        public XWPFDocument(OPCPackage pkg)
            : base(pkg)
        {
            ;

            //build a tree of POIXMLDocumentParts, this document being the root
            Load(XWPFFactory.GetInstance());
        }

        public XWPFDocument(Stream is1)
            : base(PackageHelper.Open(is1))
        {

            //build a tree of POIXMLDocumentParts, this workbook being the root
            Load(XWPFFactory.GetInstance());
        }

        public XWPFDocument()
            : base(newPackage())
        {
            onDocumentCreate();
        }


        protected void onDocumentRead()
        {
            //try {
            //    DocumentDocument doc = DocumentDocument.Factory.Parse(getPackagePart().InputStream);
            //    ctDocument = doc.Document;

            //    InitFootnotes();

            //    // parse the document with cursor and add
            //    // the XmlObject to its lists
            //    XmlCursor cursor = ctDocument.Body.NewCursor();
            //    cursor.SelectPath("./*");
            //    while (cursor.ToNextSelection()) {
            //        XmlObject o = cursor.Object;
            //        if (o is CTP) {
            //            XWPFParagraph p = new XWPFParagraph((CTP) o, this);
            //            bodyElements.Add(p);
            //            paragraphs.Add(p);
            //        } else if (o is CTTbl) {
            //            XWPFTable t = new XWPFTable((CTTbl) o, this);
            //            bodyElements.Add(t);
            //            tables.Add(t);
            //        }
            //    }
            //    cursor.Dispose();

            //    // Sort out headers and footers
            //    if (doc.Document.Body.SectPr != null)
            //        headerFooterPolicy = new XWPFHeaderFooterPolicy(this);

            //    // Create for each XML-part in the Package a PartClass
            //    foreach (POIXMLDocumentPart p in GetRelations()) {
            //        String relation = p.PackageRelationship.RelationshipType;
            //        if (relation.Equals(XWPFRelation.STYLES.Relation)) {
            //            this.styles = (XWPFStyles) p;
            //            this.styles.OnDocumentRead();
            //        } else if (relation.Equals(XWPFRelation.NUMBERING.Relation)) {
            //            this.numbering = (XWPFNumbering) p;
            //            this.numbering.OnDocumentRead();
            //        } else if (relation.Equals(XWPFRelation.FOOTER.Relation)) {
            //            XWPFFooter footer = (XWPFFooter) p;
            //            footers.Add(footer);
            //            footer.OnDocumentRead();
            //        } else if (relation.Equals(XWPFRelation.HEADER.Relation)) {
            //            XWPFHeader header = (XWPFHeader) p;
            //            headers.Add(header);
            //            header.OnDocumentRead();
            //        } else if (relation.Equals(XWPFRelation.COMMENT.Relation)) {
            //            // TODO Create according XWPFComment class, extending POIXMLDocumentPart
            //            CommentsDocument cmntdoc = CommentsDocument.Factory.Parse(p.PackagePart.InputStream);
            //            foreach (CTComment ctcomment in cmntdoc.Comments.CommentList) {
            //                comments.Add(new XWPFComment(ctcomment, this));
            //            }
            //        } else if (relation.Equals(XWPFRelation.SETTINGS.Relation)) {
            //            Settings = (XWPFSettings) p;
            //            Settings.OnDocumentRead();
            //        } else if (relation.Equals(XWPFRelation.IMAGES.Relation)) {
            //            XWPFPictureData picData = (XWPFPictureData) p;
            //            picData.OnDocumentRead();
            //            registerPackagePictureData(picData);
            //            pictures.Add(picData);
            //        }
            //    }
            //    InitHyperlinks();
            //} catch (XmlException e) {
            //    throw new POIXMLException(e);
            //}
            throw new NotImplementedException();
        }

        private void InitHyperlinks()
        {
            // Get the hyperlinks
            // TODO: make me optional/Separated in private function
            //try {
            //    Iterator<PackageRelationship> relIter =
            //        GetPackagePart().GetRelationshipsByType(XWPFRelation.HYPERLINK.Relation).iterator();
            //    while(relIter.HasNext()) {
            //        PackageRelationship rel = relIter.Next();
            //        hyperlinks.Add(new XWPFHyperlink(rel.Id, rel.TargetURI.ToString()));
            //    }
            //} catch (InvalidFormatException e){
            //    throw new POIXMLException(e);
            //}
            throw new NotImplementedException();
        }

        private void InitFootnotes()
        {
            //foreach(POIXMLDocumentPart p in GetRelations()){
            //   String relation = p.PackageRelationship.RelationshipType;
            //   if (relation.Equals(XWPFRelation.FOOTNOTE.Relation)) {
            //      FootnotesDocument footnotesDocument = FootnotesDocument.Factory.Parse(p.PackagePart.InputStream);
            //      this.footnotes = (XWPFFootnotes)p;
            //      this.footnotes.OnDocumentRead();

            //      foreach(CTFtnEdn ctFtnEdn in footnotesDocument.Footnotes.FootnoteList) {
            //         footnotes.AddFootnote(ctFtnEdn);
            //      }
            //   } else if (relation.Equals(XWPFRelation.ENDNOTE.Relation)){
            //      EndnotesDocument endnotesDocument = EndnotesDocument.Factory.Parse(p.PackagePart.InputStream);

            //      foreach(CTFtnEdn ctFtnEdn in endnotesDocument.Endnotes.EndnoteList) {
            //         endnotes.Put(ctFtnEdn.Id.IntValue(), new XWPFFootnote(this, ctFtnEdn));
            //      }
            //   }
            //}
            throw new NotImplementedException();
        }

        /**
         * Create a new WordProcessingML package and Setup the default minimal content
         */
        protected static OPCPackage newPackage()
        {
            // try {
            //    OPCPackage pkg = OPCPackage.Create(new MemoryStream());
            //    // Main part
            //    PackagePartName corePartName = PackagingURIHelper.CreatePartName(XWPFRelation.DOCUMENT.DefaultFileName);
            //    // Create main part relationship
            //    pkg.AddRelationship(corePartName, TargetMode.INTERNAL, PackageRelationshipTypes.CORE_DOCUMENT);
            //    // Create main document part
            //    pkg.CreatePart(corePartName, XWPFRelation.DOCUMENT.ContentType);

            //    pkg.PackageProperties.CreatorProperty=(DOCUMENT_CREATOR);

            //    return pkg;
            //} catch (Exception e){
            //    throw new POIXMLException(e);
            //}
            throw new NotImplementedException();
        }

        /**
         * Create a new CTWorkbook with all values Set to default
         */

        protected void onDocumentCreate()
        {
            //ctDocument = CTDocument1.Factory.NewInstance();
            //ctDocument.AddNewBody();

            //Settings = (XWPFSettings) CreateRelationship(XWPFRelation.SETTINGS,XWPFFactory.Instance);

            //POIXMLProperties.ExtendedProperties expProps = GetProperties().ExtendedProperties;
            //expProps.UnderlyingProperties.Application=(DOCUMENT_CREATOR);
            throw new NotImplementedException();
        }

        /**
         * Returns the low level document base object
         */
        //CTDocument1
        public CT_Document GetDocument()
        {
            //return ctDocument;
            throw new NotImplementedException();
        }

        IdentifierManager GetDrawingIdManager()
        {
            return DrawingIdManager;
        }

        /**
         * returns an Iterator with paragraphs and tables
         * @see NPOI.XWPF.UserModel.IBody#getBodyElements()
         */
        public List<IBodyElement> GetBodyElements()
        {
            //return Collections.UnmodifiableList(bodyElements);
            throw new NotImplementedException();
        }

        /**
         * @see NPOI.XWPF.UserModel.IBody#getParagraphs()
         */
        public List<XWPFParagraph> GetParagraphs()
        {
            //return Collections.UnmodifiableList(paragraphs);
            throw new NotImplementedException();
        }

        /**
         * @see NPOI.XWPF.UserModel.IBody#getTables()
         */
        public List<XWPFTable> GetTables()
        {
            //return Collections.UnmodifiableList(tables);
            throw new NotImplementedException();
        }

        /**
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
         * 
         * @return  the list of footers
         */
        public List<XWPFFooter> GetFooterList()
        {
            //return Collections.UnmodifiableList(footers);
            throw new NotImplementedException();
        }

        public XWPFFooter GetFooterArray(int pos)
        {
            return footers[(pos)];
        }

        /**
         * 
         * @return  the list of headers
         */
        public List<XWPFHeader> GetHeaderList()
        {
            //return Collections.UnmodifiableList(headers);
            throw new NotImplementedException();
        }

        public XWPFHeader GetHeaderArray(int pos)
        {
            return headers[(pos)];
        }

        public String GetTblStyle(XWPFTable table)
        {
            //return table.StyleID;
            throw new NotImplementedException();
        }

        public XWPFHyperlink GetHyperlinkByID(String id)
        {
            //Iterator<XWPFHyperlink> iter = hyperlinks.Iterator();
            //while (iter.HasNext()) {
            //    XWPFHyperlink link = iter.Next();
            //    if(link.Id.Equals(id))
            //        return link;
            //}

            //return null;
            throw new NotImplementedException();
        }

        public XWPFFootnote GetFootnoteByID(int id)
        {
            return footnotes.GetFootnoteById(id);
        }

        public XWPFFootnote GetEndnoteByID(int id)
        {
            return endnotes[(id)];
        }

        public List<XWPFFootnote> GetFootnotes()
        {
            //return footnotes.FootnotesList;
            throw new NotImplementedException();
        }

        public XWPFHyperlink[] GetHyperlinks()
        {
            //return hyperlinks.ToArray(new XWPFHyperlink[hyperlinks.Size()]);
            throw new NotImplementedException();
        }

        public XWPFComment GetCommentByID(String id)
        {
            //Iterator<XWPFComment> iter = comments.Iterator();
            //while (iter.HasNext()) {
            //    XWPFComment comment = iter.Next();
            //    if(comment.Id.Equals(id))
            //        return comment;
            //}

            //return null;
            throw new NotImplementedException();
        }

        public XWPFComment[] GetComments()
        {
            //return comments.ToArray(new XWPFComment[comments.Size()]);
            throw new NotImplementedException();
        }

        /**
         * Get the document part that's defined as the
         *  given relationship of the core document.
         */
        public PackagePart GetPartById(String id)
        {
            //try {
            //    return GetTargetPart(getCorePart().GetRelationship(id));
            //} catch (InvalidFormatException e) {
            //    throw new ArgumentException(e);
            //}
            throw new NotImplementedException();
        }

        /**
         * Returns the policy on headers and footers, which
         *  also provides a way to Get at them.
         */
        public XWPFHeaderFooterPolicy GetHeaderFooterPolicy()
        {
            //return headerFooterPolicy;
            throw new NotImplementedException();
        }

        /**
         * Returns the styles object used
         */

        public CT_Styles GetStyle()
        {
            //PackagePart[] parts;
            //try {
            //    parts = GetRelatedByType(XWPFRelation.STYLES.Relation);
            //} catch(InvalidFormatException e) {
            //    throw new InvalidOperationException(e);
            //}
            //if(parts.Length != 1) {
            //    throw new InvalidOperationException("Expecting one Styles document part, but found " + parts.Length);
            //}

            //StylesDocument sd = StylesDocument.Factory.Parse(parts[0].InputStream);
            //return sd.Styles;
            throw new NotImplementedException();
        }

        /**
         * Get the document's embedded files.
         */

        public override List<PackagePart> GetAllEmbedds()
        {
            //List<PackagePart> embedds = new LinkedList<PackagePart>();

            //// Get the embeddings for the workbook
            //foreach (PackageRelationship rel in GetPackagePart().GetRelationshipsByType(OLE_OBJECT_REL_TYPE)) {
            //    embedds.Add(getTargetPart(rel));
            //}

            //foreach (PackageRelationship rel in GetPackagePart().GetRelationshipsByType(PACK_OBJECT_REL_TYPE)) {
            //    embedds.Add(getTargetPart(rel));
            //}

            //return embedds;
            throw new NotImplementedException();
        }

        /**
         * Finds that for example the 2nd entry in the body list is the 1st paragraph
         */
        private int GetBodyElementSpecificPos(int pos, List<IBodyElement> list)
        {
            // If there's nothing to Find, skip it
            //if(list.Size() == 0) {
            //   return -1;
            //}

            //if(pos >= 0 && pos < bodyElements.Size()) {
            //   // Ensure the type is correct
            //   IBodyElement needle = bodyElements.Get(pos);
            //   if(needle.ElementType != list.Get(0).ElementType) {
            //      // Wrong type
            //      return -1;
            //   }

            //   // Work back until we find it
            //   int startPos = Math.Min(pos, list.Size()-1);
            //   for(int i=startPos; i>=0; i--) {
            //      if(list.Get(i) == needle) {
            //         return i;
            //      }
            //   }
            //}

            //// Couldn't be found
            //return -1;
            throw new NotImplementedException();
        }

        /**
         * Look up the paragraph at the specified position in the body elemnts list
         * and return this paragraphs position in the paragraphs list
         * 
         * @param pos
         *            The position of the relevant paragraph in the body elements
         *            list
         * @return the position of the paragraph in the paragraphs list, if there is
         *         a paragraph at the position in the bodyelements list. Else it
         *         will return -1
         * 
         */
        public int GetParagraphPos(int pos)
        {
            //return GetBodyElementSpecificPos(pos, paragraphs);
            throw new NotImplementedException();
        }

        /**
         * Get with the position of a table in the bodyelement array list 
         * the position of this table in the table array list
         * @param pos position of the table in the bodyelement array list
         * @return if there is a table at the position in the bodyelement array list,
         * 		   else it will return null. 
         */
        public int GetTablePos(int pos)
        {
            //return GetBodyElementSpecificPos(pos, tables);
            throw new NotImplementedException();
        }

        /**
         * Add a new paragraph at position of the cursor. The cursor must be on the
         * {@link TokenType#START} tag of an subelement of the documents body. When
         * this method is done, the cursor passed as parameter points to the
         * {@link TokenType#END} of the newly inserted paragraph.
         * 
         * @param cursor
         * @return the {@link XWPFParagraph} object representing the newly inserted
         *         CTP object
         */
        public XWPFParagraph insertNewParagraph(/*XmlCursor*/XmlDocument cursor)
        {
            //if (isCursorInBody(cursor)) {
            //    String uri = CTP.type.Name.NamespaceURI;
            //    /*
            //     * TODO DO not use a coded constant, find the constant in the OOXML
            //     * classes instead, as the child of type CT_Paragraph is defined in the 
            //     * OOXML schema as 'p'
            //     */
            //    String localPart = "p";
            //    // Creates a new Paragraph, cursor is positioned inside the new
            //    // element
            //    cursor.BeginElement(localPart, uri);
            //    // Move the cursor to the START token to the paragraph just Created
            //    cursor.ToParent();
            //    CTP p = (CTP) cursor.Object;
            //    XWPFParagraph newP = new XWPFParagraph(p, this);
            //    XmlObject o = null;
            //    /*
            //     * Move the cursor to the previous element until a) the next
            //     * paragraph is found or b) all elements have been passed
            //     */
            //    while (!(o is CTP) && (cursor.ToPrevSibling())) {
            //        o = cursor.Object;
            //    }
            //    /*
            //     * if the object that has been found is a) not a paragraph or b) is
            //     * the paragraph that has just been inserted, as the cursor in the
            //     * while loop above was not Moved as there were no other siblings,
            //     * then the paragraph that was just inserted is the first paragraph
            //     * in the body. Otherwise, take the previous paragraph and calculate
            //     * the new index for the new paragraph.
            //     */
            //    if ((!(o is CTP)) || (CTP) o == p) {
            //        paragraphs.Add(0, newP);
            //    } else {
            //        int pos = paragraphs.IndexOf(getParagraph((CTP) o)) + 1;
            //        paragraphs.Add(pos, newP);
            //    }

            //    /*
            //     * create a new cursor, that points to the START token of the just
            //     * inserted paragraph
            //     */
            //    XmlCursor newParaPos = p.NewCursor();
            //    try {
            //        /*
            //         * Calculate the paragraphs index in the list of all body
            //         * elements
            //         */
            //        int i = 0;
            //        cursor.ToCursor(newParaPos);
            //        while (cursor.ToPrevSibling()) {
            //            o = cursor.Object;
            //            if (o is CTP || o is CTTbl)
            //                i++;
            //        }
            //        bodyElements.Add(i, newP);
            //        cursor.ToCursor(newParaPos);
            //        cursor.ToEndToken();
            //        return newP;
            //    } finally {
            //        newParaPos.Dispose();
            //    }
            //}
            //return null;
            throw new NotImplementedException();
        }

        public XWPFTable insertNewTbl(/*XmlCursor*/XmlDocument cursor)
        {
            //if (isCursorInBody(cursor)) {
            //    String uri = CTTbl.type.Name.NamespaceURI;
            //    String localPart = "tbl";
            //    cursor.BeginElement(localPart, uri);
            //    cursor.ToParent();
            //    CTTbl t = (CTTbl) cursor.Object;
            //    XWPFTable newT = new XWPFTable(t, this);
            //    cursor.RemoveXmlContents();
            //    XmlObject o = null;
            //    while (!(o is CTTbl) && (cursor.ToPrevSibling())) {
            //        o = cursor.Object;
            //    }
            //    if (!(o is CTTbl)) {
            //        tables.Add(0, newT);
            //    } else {
            //        int pos = tables.IndexOf(getTable((CTTbl) o)) + 1;
            //        tables.Add(pos, newT);
            //    }
            //    int i = 0;
            //    cursor = t.NewCursor();
            //    while (cursor.ToPrevSibling()) {
            //        o = cursor.Object;
            //        if (o is CTP || o is CTTbl)
            //            i++;
            //    }
            //    bodyElements.Add(i, newT);
            //    cursor = t.NewCursor();
            //    cursor.ToEndToken();
            //    return newT;
            //}
            //return null;
            throw new NotImplementedException();
        }

        /**
         * verifies that cursor is on the right position
         * @param cursor
         */
        private bool IsCursorInBody(/*XmlCursor*/XmlDocument cursor)
        {
            /*XmlCursor verify = cursor.NewCursor();
            verify.ToParent();
            try {
                return (verify.Object == this.ctDocument.Body);
            } finally {
                verify.Dispose();
            }*/
            throw new NotImplementedException();
        }

        private int GetPosOfBodyElement(IBodyElement needle)
        {
            /* BodyElementType type = needle.ElementType;
            IBodyElement current; 
             for(int i=0; i<bodyElements.Size(); i++) {
                current = bodyElements.Get(i);
                if(current.ElementType == type) {
                   if(current.Equals(needle)) {
                      return i;
                   }
                }
             }
             return -1;*/
            throw new NotImplementedException();
        }

        /**
         * Get the position of the paragraph, within the list
         *  of all the body elements.
         * @param p The paragraph to find
         * @return The location, or -1 if the paragraph couldn't be found 
         */
        public int GetPosOfParagraph(XWPFParagraph p)
        {
            return GetPosOfBodyElement(p);
        }

        /**
         * Get the position of the table, within the list of
         *  all the body elements.
         * @param t The table to find
         * @return The location, or -1 if the table couldn't be found
         */
        public int GetPosOfTable(XWPFTable t)
        {
            return GetPosOfBodyElement(t);
        }

        /**
         * Commit and saves the document
         */

        protected void Commit()
        {

            //XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            //xmlOptions.SaveSyntheticDocumentElement=(new QName(CTDocument1.type.Name.NamespaceURI, "document"));
            //Dictionary<String, String> map = new Dictionary<String, String>();
            //map.Put("http://schemas.Openxmlformats.org/officeDocument/2006/math", "m");
            //map.Put("urn:schemas-microsoft-com:office:office", "o");
            //map.Put("http://schemas.Openxmlformats.org/officeDocument/2006/relationships", "r");
            //map.Put("urn:schemas-microsoft-com:vml", "v");
            //map.Put("http://schemas.Openxmlformats.org/markup-compatibility/2006", "ve");
            //map.Put("http://schemas.Openxmlformats.org/wordProcessingml/2006/main", "w");
            //map.Put("urn:schemas-microsoft-com:office:word", "w10");
            //map.Put("http://schemas.microsoft.com/office/word/2006/wordml", "wne");
            //map.Put("http://schemas.Openxmlformats.org/drawingml/2006/wordProcessingDrawing", "wp");
            //xmlOptions.SaveSuggestedPrefixes=(map);

            //PackagePart part = GetPackagePart();
            //OutputStream out1 = part.OutputStream;
            //ctDocument.Save(out, xmlOptions);
            //out1.Close();
            throw new NotImplementedException();
        }

        /**
         * Gets the index of the relation we're trying to create
         * @param relation
         * @return i
         */
        private int GetRelationIndex(XWPFRelation relation)
        {
            //List<POIXMLDocumentPart> relations = GetRelations();
            //int i = 1;
            //for (Iterator<POIXMLDocumentPart> it = relations.Iterator(); it.HasNext() ; ) {
            //   POIXMLDocumentPart item = it.Next();
            //   if (item.PackageRelationship.RelationshipType.Equals(relation.Relation)) {
            //      i++;
            //   }
            //}
            //return i;
            throw new NotImplementedException();
        }

        /**
         * Appends a new paragraph to this document
         * @return a new paragraph
         */
        public XWPFParagraph CreateParagraph()
        {
            //XWPFParagraph p = new XWPFParagraph(ctDocument.Body.AddNewP(), this);
            //bodyElements.Add(p);
            //paragraphs.Add(p);
            //return p;
            throw new NotImplementedException();
        }

        /**
         * Creates an empty numbering if one does not already exist and Sets the numbering member
         * @return numbering
         */
        public XWPFNumbering CreateNumbering()
        {
            //if(numbering == null) {
            //    NumberingDocument numberingDoc = NumberingDocument.Factory.NewInstance();

            //    XWPFRelation relation = XWPFRelation.NUMBERING;
            //    int i = GetRelationIndex(relation);

            //    XWPFNumbering wrapper = (XWPFNumbering)CreateRelationship(relation, XWPFFactory.Instance, i);
            //    wrapper.Numbering=(numberingDoc.AddNewNumbering());
            //    numbering = wrapper;
            //}

            //return numbering;
            throw new NotImplementedException();
        }

        /**
         * Creates an empty styles for the document if one does not already exist
         * @return styles
         */
        public XWPFStyles CreateStyles()
        {
            //if(styles == null) {
            //   StylesDocument stylesDoc = StylesDocument.Factory.NewInstance();

            //   XWPFRelation relation = XWPFRelation.STYLES;
            //   int i = GetRelationIndex(relation);

            //   XWPFStyles wrapper = (XWPFStyles)CreateRelationship(relation, XWPFFactory.Instance, i);
            //   wrapper.Styles=(stylesDoc.AddNewStyles());
            //   styles = wrapper;
            //}

            //return styles;
            throw new NotImplementedException();
        }

        /**
         * Creates an empty footnotes element for the document if one does not already exist
         * @return footnotes
         */
        public XWPFFootnotes CreateFootnotes()
        {
            //if(footnotes == null) {
            //   FootnotesDocument footnotesDoc = FootnotesDocument.Factory.NewInstance();

            //   XWPFRelation relation = XWPFRelation.FOOTNOTE;
            //   int i = GetRelationIndex(relation);

            //   XWPFFootnotes wrapper = (XWPFFootnotes)CreateRelationship(relation, XWPFFactory.Instance, i);
            //   wrapper.Footnotes=(footnotesDoc.AddNewFootnotes());
            //   footnotes = wrapper;
            //}

            //return footnotes;
            throw new NotImplementedException();
        }

        public XWPFFootnote AddFootnote(CT_FtnEdn note)
        {
            return footnotes.AddFootnote(note);
        }

        public XWPFFootnote AddEndnote(CT_FtnEdn note)
        {
            //XWPFFootnote endnote = new XWPFFootnote(this, note); 
            //endnotes.Put(note.Id.IntValue(), endnote);
            //return endnote;
            throw new NotImplementedException();
        }

        /**
         * remove a BodyElement from bodyElements array list 
         * @param pos
         * @return true if removing was successfully, else return false
         */
        public bool RemoveBodyElement(int pos)
        {
            //if(pos >= 0 && pos < bodyElements.Size()) {
            //   BodyElementType type = bodyElements.Get(pos).ElementType; 
            //    if(type == BodyElementType.TABLE){
            //        int tablePos = GetTablePos(pos);
            //        tables.Remove(tablePos);
            //        ctDocument.Body.RemoveTbl(tablePos);
            //    }
            //    if(type == BodyElementType.PARAGRAPH){
            //        int paraPos = GetParagraphPos(pos);
            //        paragraphs.Remove(paraPos);
            //        ctDocument.Body.RemoveP(paraPos);
            //    }
            // bodyElements.Remove(pos);
            // return true;            
            //}
            //return false;
            throw new NotImplementedException();
        }

        /**
         * copies content of a paragraph to a existing paragraph in the list paragraphs at position pos
         * @param paragraph
         * @param pos
         */
        public void SetParagraph(XWPFParagraph paragraph, int pos)
        {
            //paragraphs.Set(pos, paragraph);
            //ctDocument.Body.PArray=(pos, paragraph.CTP);
            /* TODO update body element, update xwpf element, verify that
             * incoming paragraph belongs to this document or if not, XML was
             * copied properly (namespace-abbreviations, etc.)
             */
            throw new NotImplementedException();
        }

        /**
         * @return the LastParagraph of the document
         */
        public XWPFParagraph GetLastParagraph()
        {
            int lastPos = paragraphs.ToArray().Length - 1;
            return paragraphs[(lastPos)];
        }

        /**
         * Create an empty table with one row and one column as default.
         * @return a new table
         */
        public XWPFTable CreateTable()
        {
            //XWPFTable table = new XWPFTable(ctDocument.Body.AddNewTbl(), this);
            //bodyElements.Add(table);
            //tables.Add(table);
            //return table;
            throw new NotImplementedException();
        }

        /**
         * Create an empty table with a number of rows and cols specified
         * @param rows
         * @param cols
         * @return table
         */
        public XWPFTable CreateTable(int rows, int cols)
        {
            //XWPFTable table = new XWPFTable(ctDocument.Body.AddNewTbl(), this, rows, cols);
            //bodyElements.Add(table);
            //tables.Add(table);
            //return table;
            throw new NotImplementedException();
        }

        /**
         * 
         */
        public void CreateTOC()
        {
            //CTSdtBlock block = this.Document.Body.AddNewSdt();
            //TOC toc = new TOC(block);
            //foreach (XWPFParagraph par in paragraphs) {
            //    String parStyle = par.Style;
            //    if (parStyle != null && parStyle.Substring(0, 7).Equals("Heading")) {
            //        try {
            //            int level = Int32.ValueOf(parStyle.Substring("Heading".Length)).intValue();
            //            toc.AddRow(level, par.Text, 1, "112723803");
            //        } catch (FormatException e) {
            //            e.PrintStackTrace();
            //        }
            //    }
            //}
            throw new NotImplementedException();
        }

        /**Replace content of table in array tables at position pos with a
         * @param pos
         * @param table
         */
        public void SetTable(int pos, XWPFTable table)
        {
            //tables.Set(pos, table);
            //ctDocument.Body.TblArray=(pos, table.CTTbl);
            throw new NotImplementedException();
        }

        /**
         * Verifies that the documentProtection tag in Settings.xml file <br/>
         * specifies that the protection is enforced (w:enforcement="1") <br/>
         * and that the kind of protection is ReadOnly (w:edit="readOnly")<br/>
         * <br/>
         * sample snippet from Settings.xml
         * <pre>
         *     &lt;w:settings  ... &gt;
         *         &lt;w:documentProtection w:edit=&quot;readOnly&quot; w:enforcement=&quot;1&quot;/&gt;
         * </pre>
         * 
         * @return true if documentProtection is enforced with option ReadOnly
         */
        public bool IsEnforcedReadonlyProtection()
        {
            //return Settings.IsEnforcedWith(STDocProtect.READ_ONLY);
            throw new NotImplementedException();
        }

        /**
         * Verifies that the documentProtection tag in Settings.xml file <br/>
         * specifies that the protection is enforced (w:enforcement="1") <br/>
         * and that the kind of protection is forms (w:edit="forms")<br/>
         * <br/>
         * sample snippet from Settings.xml
         * <pre>
         *     &lt;w:settings  ... &gt;
         *         &lt;w:documentProtection w:edit=&quot;forms&quot; w:enforcement=&quot;1&quot;/&gt;
         * </pre>
         * 
         * @return true if documentProtection is enforced with option forms
         */
        public bool IsEnforcedFillingFormsProtection()
        {
            //return Settings.IsEnforcedWith(STDocProtect.FORMS);
            throw new NotImplementedException();
        }

        /**
         * Verifies that the documentProtection tag in Settings.xml file <br/>
         * specifies that the protection is enforced (w:enforcement="1") <br/>
         * and that the kind of protection is comments (w:edit="comments")<br/>
         * <br/>
         * sample snippet from Settings.xml
         * <pre>
         *     &lt;w:settings  ... &gt;
         *         &lt;w:documentProtection w:edit=&quot;comments&quot; w:enforcement=&quot;1&quot;/&gt;
         * </pre>
         * 
         * @return true if documentProtection is enforced with option comments
         */
        public bool IsEnforcedCommentsProtection()
        {
            //return Settings.IsEnforcedWith(STDocProtect.COMMENTS);
            throw new NotImplementedException();
        }

        /**
         * Verifies that the documentProtection tag in Settings.xml file <br/>
         * specifies that the protection is enforced (w:enforcement="1") <br/>
         * and that the kind of protection is trackedChanges (w:edit="trackedChanges")<br/>
         * <br/>
         * sample snippet from Settings.xml
         * <pre>
         *     &lt;w:settings  ... &gt;
         *         &lt;w:documentProtection w:edit=&quot;trackedChanges&quot; w:enforcement=&quot;1&quot;/&gt;
         * </pre>
         * 
         * @return true if documentProtection is enforced with option trackedChanges
         */
        public bool IsEnforcedTrackedChangesProtection()
        {
            //return Settings.IsEnforcedWith(STDocProtect.TRACKED_CHANGES);
            throw new NotImplementedException();
        }

        /**
         * Enforces the ReadOnly protection.<br/>
         * In the documentProtection tag inside Settings.xml file, <br/>
         * it Sets the value of enforcement to "1" (w:enforcement="1") <br/>
         * and the value of edit to ReadOnly (w:edit="readOnly")<br/>
         * <br/>
         * sample snippet from Settings.xml
         * <pre>
         *     &lt;w:settings  ... &gt;
         *         &lt;w:documentProtection w:edit=&quot;readOnly&quot; w:enforcement=&quot;1&quot;/&gt;
         * </pre>
         */
        public void enforceReadonlyProtection()
        {
            //Settings.EnforcementEditValue=(STDocProtect.READ_ONLY);
            throw new NotImplementedException();
        }

        /**
         * Enforce the Filling Forms protection.<br/>
         * In the documentProtection tag inside Settings.xml file, <br/>
         * it Sets the value of enforcement to "1" (w:enforcement="1") <br/>
         * and the value of edit to forms (w:edit="forms")<br/>
         * <br/>
         * sample snippet from Settings.xml
         * <pre>
         *     &lt;w:settings  ... &gt;
         *         &lt;w:documentProtection w:edit=&quot;forms&quot; w:enforcement=&quot;1&quot;/&gt;
         * </pre>
         */
        public void enforceFillingFormsProtection()
        {
            //Settings.EnforcementEditValue=(STDocProtect.FORMS);
            throw new NotImplementedException();
        }

        /**
         * Enforce the Comments protection.<br/>
         * In the documentProtection tag inside Settings.xml file,<br/>
         * it Sets the value of enforcement to "1" (w:enforcement="1") <br/>
         * and the value of edit to comments (w:edit="comments")<br/>
         * <br/>
         * sample snippet from Settings.xml
         * <pre>
         *     &lt;w:settings  ... &gt;
         *         &lt;w:documentProtection w:edit=&quot;comments&quot; w:enforcement=&quot;1&quot;/&gt;
         * </pre>
         */
        public void enforceCommentsProtection()
        {
            //Settings.EnforcementEditValue=(STDocProtect.COMMENTS);
            throw new NotImplementedException();
        }

        /**
         * Enforce the Tracked Changes protection.<br/>
         * In the documentProtection tag inside Settings.xml file, <br/>
         * it Sets the value of enforcement to "1" (w:enforcement="1") <br/>
         * and the value of edit to trackedChanges (w:edit="trackedChanges")<br/>
         * <br/>
         * sample snippet from Settings.xml
         * <pre>
         *     &lt;w:settings  ... &gt;
         *         &lt;w:documentProtection w:edit=&quot;trackedChanges&quot; w:enforcement=&quot;1&quot;/&gt;
         * </pre>
         */
        public void enforceTrackedChangesProtection()
        {
            //Settings.EnforcementEditValue=(STDocProtect.TRACKED_CHANGES);
            throw new NotImplementedException();
        }

        /**
         * Remove protection enforcement.<br/>
         * In the documentProtection tag inside Settings.xml file <br/>
         * it Sets the value of enforcement to "0" (w:enforcement="0") <br/>
         */
        public void RemoveProtectionEnforcement()
        {
            Settings.RemoveEnforcement();
        }

        /**
         * inserts an existing XWPFTable to the arrays bodyElements and tables
         * @param pos
         * @param table
         */
        public void insertTable(int pos, XWPFTable table)
        {
            //bodyElements.Add(pos, table);
            //int i;
            //for (i = 0; i < ctDocument.Body.TblList.Size(); i++) {
            //    CTTbl tbl = ctDocument.Body.GetTblArray(i);
            //    if (tbl == table.CTTbl) {
            //        break;
            //    }
            //}
            //tables.Add(i, table);
            throw new NotImplementedException();
        }

        /**
         * Returns all Pictures, which are referenced from the document itself.
         * @return a {@link List} of {@link XWPFPictureData}. The returned {@link List} is unmodifiable. Use #a
         */
        public List<XWPFPictureData> GetAllPictures()
        {
            //return Collections.UnmodifiableList(pictures);
            throw new NotImplementedException();
        }

        /**
         * @return all Pictures in this package
         */
        public List<XWPFPictureData> GetAllPackagePictures()
        {
            //List<XWPFPictureData> result = new List<XWPFPictureData>();
            //Collection<List<XWPFPictureData>> values = packagePictures.Values();
            //foreach (List<XWPFPictureData> list in values) {
            //    result.AddAll(list);
            //}
            //return Collections.UnmodifiableList(result);
            throw new NotImplementedException();
        }

        void registerPackagePictureData(XWPFPictureData picData)
        {
            //List<XWPFPictureData> list = packagePictures.Get(picData.Checksum);
            //if (list == null) {
            //    list = new List<XWPFPictureData>(1);
            //    packagePictures.Put(picData.Checksum, list);
            //}
            //if (!list.Contains(picData))
            //{
            //    list.Add(picData);
            //}
            throw new NotImplementedException();
        }

        XWPFPictureData FindPackagePictureData(byte[] pictureData, int format)
        {
            //long Checksum = IOUtils.CalculateChecksum(pictureData);
            //XWPFPictureData xwpfPicData = null;
            ///*
            // * Try to find PictureData with this Checksum. Create new, if none
            // * exists.
            // */
            //List<XWPFPictureData> xwpfPicDataList = packagePictures.Get(Checksum);
            //if (xwpfPicDataList != null) {
            //    Iterator<XWPFPictureData> iter = xwpfPicDataList.Iterator();
            //    while (iter.HasNext() && xwpfPicData == null) {
            //        XWPFPictureData curElem = iter.Next();
            //        if (Arrays.Equals(pictureData, curElem.Data)) {
            //            xwpfPicData = curElem;
            //        }
            //    }
            //} 
            //return xwpfPicData;
            throw new NotImplementedException();
        }

        public String AddPictureData(byte[] pictureData, int format)
        {
            XWPFPictureData xwpfPicData = FindPackagePictureData(pictureData, format);
            POIXMLRelation relDesc = XWPFPictureData.RELATIONS[format];

            //if (xwpfPicData == null)
            //{
            //    /* Part doesn't exist, create a new one */
            //    int idx = GetNextPicNameNumber(format);
            //    xwpfPicData = (XWPFPictureData) CreateRelationship(relDesc, XWPFFactory.Instance,idx);
            //    /* write bytes to new part */
            //    PackagePart picDataPart = xwpfPicData.PackagePart;
            //    OutputStream out1 = null;
            //    try {
            //        out1 = picDataPart.OutputStream;
            //        out1.Write(pictureData);
            //    } catch (IOException e) {
            //        throw new POIXMLException(e);
            //    } finally {
            //        try {
            //            out1.Close();
            //        } catch (IOException e) {
            //            // ignore
            //        }
            //    }

            //    registerPackagePictureData(xwpfPicData);
            //    pictures.Add(xwpfPicData);

            //    return GetRelationId(xwpfPicData);
            //}
            //else if (!getRelations().Contains(xwpfPicData))
            //{
            //    /*
            //     * Part already existed, but was not related so far. Create
            //     * relationship to the already existing part and update
            //     * POIXMLDocumentPart data.
            //     */
            //    PackagePart picDataPart = xwpfPicData.PackagePart;
            //    // TODO add support for TargetMode.EXTERNAL relations.
            //    TargetMode targetMode = TargetMode.INTERNAL;
            //    PackagePartName partName = picDataPart.PartName;
            //    String relation = relDesc.Relation;
            //    PackageRelationship relShip = GetPackagePart().AddRelationship(partName,targetMode,relation);
            //    String id = relShip.Id;
            //    AddRelation(id,xwpfPicData);
            //    pictures.Add(xwpfPicData);
            //    return id;
            //}
            //else 
            //{
            //    /* Part already existed, Get relation id and return it */
            //    return GetRelationId(xwpfPicData);
            //}
            throw new NotImplementedException();
        }

        public String AddPictureData(Stream is1, int format)
        {
            try
            {
                byte[] data = IOUtils.ToByteArray(is1);
                return AddPictureData(data, format);
            }
            catch (IOException e)
            {
                throw new POIXMLException(e);
            }
        }

        /**
         * Get the next free ImageNumber
         * @param format
         * @return the next free ImageNumber
         * @throws InvalidFormatException 
         */
        public int GetNextPicNameNumber(int format)
        {
            /*int img = GetAllPackagePictures().Count + 1;
            String proposal = XWPFPictureData.RELATIONS[format].GetFileName(img);
            PackagePartName CreatePartName = PackagingURIHelper.CreatePartName(proposal);
            while (this.Package.GetPart(CreatePartName) != null) {
                img++;
                proposal = XWPFPictureData.RELATIONS[format].GetFileName(img);
                CreatePartName = PackagingURIHelper.CreatePartName(proposal);
            }
            return img;*/
            throw new NotImplementedException();
        }

        /**
         * returns the PictureData by blipID
         * @param blipID
         * @return XWPFPictureData of a specificID
         */
        public XWPFPictureData GetPictureDataByID(String blipID)
        {
            POIXMLDocumentPart relatedPart = GetRelationById(blipID);
            if (relatedPart is XWPFPictureData)
            {
                XWPFPictureData xwpfPicData = (XWPFPictureData)relatedPart;
                return xwpfPicData;
            }
            return null;
        }

        /**
         * GetNumbering
         * @return numbering
         */
        public XWPFNumbering GetNumbering()
        {
            return numbering;
        }

        /**
         * Get Styles
         * @return styles for this document
         */
        public XWPFStyles GetStyles()
        {
            return styles;
        }

        /**
         * Get the paragraph with the CTP class p
         * 
         * @param p
         * @return the paragraph with the CTP class p
         */
        public XWPFParagraph GetParagraph(CT_P p)
        {
            //for (int i = 0; i < GetParagraphs().Count; i++) {
            //    if (getParagraphs().Get(i).CTP == p) {
            //        return GetParagraphs().Get(i);
            //    }
            //}
            //return null;
            throw new NotImplementedException();
        }

        /**
         * Get a table by its CTTbl-Object
         * @param ctTbl
         * @see NPOI.XWPF.UserModel.IBody#getTable(org.Openxmlformats.schemas.wordProcessingml.x2006.main.CTTbl)
         * @return a table by its CTTbl-Object or null
         */
        public XWPFTable GetTable(CT_Tbl ctTbl)
        {
            //for (int i = 0; i < tables.Size(); i++) {
            //    if (getTables().Get(i).CTTbl == ctTbl) {
            //        return GetTables().Get(i);
            //    }
            //}
            //return null;
            throw new NotImplementedException();
        }


        public IEnumerator<XWPFTable> GetTablesEnumerator()
        {
            return tables.GetEnumerator();
        }

        public IEnumerator<XWPFParagraph> GetParagraphsEnumerator()
        {
            return paragraphs.GetEnumerator();
        }

        /**
         * Returns the paragraph that of position pos
         * @see NPOI.XWPF.UserModel.IBody#getParagraphArray(int)
         */
        public XWPFParagraph GetParagraphArray(int pos)
        {
            if (pos >= 0 && pos < paragraphs.Count)
            {
                return paragraphs[(pos)];
            }
            return null;
        }

        /**
         * returns the Part, to which the body belongs, which you need for Adding relationship to other parts
         * Actually it is needed of the class XWPFTableCell. Because you have to know to which part the tableCell
         * belongs.
         * @see NPOI.XWPF.UserModel.IBody#getPart()
         */
        public POIXMLDocumentPart GetPart()
        {
            return this;
        }


        /**
         * Get the PartType of the body, for example
         * DOCUMENT, HEADER, FOOTER,	FOOTNOTE,
         *
         * @see NPOI.XWPF.UserModel.IBody#getPartType()
         */
        public BodyType GetPartType()
        {
            return BodyType.DOCUMENT;
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
            return tableRow.GetTableCell(cell);
             */
            throw new NotImplementedException();
        }

        public XWPFDocument GetXWPFDocument()
        {
            return this;
        }
    } // end class
}