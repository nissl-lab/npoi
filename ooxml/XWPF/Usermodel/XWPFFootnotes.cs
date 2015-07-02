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
    using NPOI.OpenXml4Net.OPC;
    using System.IO;
    using System.Xml;
    using System.Xml.Serialization;


    /**
     * Looks After the collection of Footnotes for a document
     *  
     * @author Mike McEuen (mceuen@hp.com)
     */
    public class XWPFFootnotes : POIXMLDocumentPart
    {
        private List<XWPFFootnote> listFootnote = new List<XWPFFootnote>();
        private CT_Footnotes ctFootnotes;

        protected XWPFDocument document;

        /**
         * Construct XWPFFootnotes from a package part
         *
         * @param part the package part holding the data of the footnotes,
         * @param rel  the package relationship of type "http://schemas.Openxmlformats.org/officeDocument/2006/relationships/footnotes"
         */
        public XWPFFootnotes(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {
            ;
        }

        /**
         * Construct XWPFFootnotes from scratch for a new document.
         */
        public XWPFFootnotes()
        {
        }

        /**
         * Read document
         */

        internal override void OnDocumentRead()
        {
            FootnotesDocument notesDoc;
            try {
               XmlDocument xmldoc = ConvertStreamToXml(GetPackagePart().GetInputStream());
               notesDoc = FootnotesDocument.Parse(xmldoc, NamespaceManager);
               ctFootnotes = notesDoc.Footnotes;
            } catch (XmlException) {
               throw new POIXMLException();
            }
	   
            //get any Footnote
            if (ctFootnotes.footnote != null)
            {
                foreach (CT_FtnEdn note in ctFootnotes.footnote)
                {
                    listFootnote.Add(new XWPFFootnote(note, this));
                }
            }
        }


        protected internal override void Commit()
        {
            /*XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            xmlOptions.SaveSyntheticDocumentElement=(new QName(CTFootnotes.type.Name.NamespaceURI, "footnotes"));
            Dictionary<String,String> map = new Dictionary<String,String>();
            map.Put("http://schemas.Openxmlformats.org/officeDocument/2006/relationships", "r");
            map.Put("http://schemas.Openxmlformats.org/wordProcessingml/2006/main", "w");
            xmlOptions.SaveSuggestedPrefixes=(map);*/
            PackagePart part = GetPackagePart();
            using (Stream out1 = part.GetOutputStream())
            {
                FootnotesDocument notesDoc = new FootnotesDocument(ctFootnotes);
                notesDoc.Save(out1);
            }
        }

        public List<XWPFFootnote> GetFootnotesList()
        {
            return listFootnote;
        }

        public XWPFFootnote GetFootnoteById(int id)
        {
            foreach(XWPFFootnote note in listFootnote) {
                if(note.GetCTFtnEdn().id == id.ToString())
                    return note;
            }
            return null;
        }

        /**
         * Sets the ctFootnotes
         * @param footnotes
         */
        public void SetFootnotes(CT_Footnotes footnotes)
        {
            ctFootnotes = footnotes;
        }

        /**
         * add an XWPFFootnote to the document
         * @param footnote
         * @throws IOException		 
         */
        public void AddFootnote(XWPFFootnote footnote)
        {
            listFootnote.Add(footnote);
            ctFootnotes.AddNewFootnote().Set(footnote.GetCTFtnEdn());
        }

        /**
         * add a footnote to the document
         * @param note
         * @throws IOException		 
         */
        public XWPFFootnote AddFootnote(CT_FtnEdn note)
        {
            CT_FtnEdn newNote = ctFootnotes.AddNewFootnote();
            newNote.Set(note);
            XWPFFootnote xNote = new XWPFFootnote(newNote, this);
            listFootnote.Add(xNote);
            return xNote;
        }

        public void SetXWPFDocument(XWPFDocument doc)
        {
            document = doc;
        }

        /**
         * @see NPOI.XWPF.UserModel.IBody#getPart()
         */
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
    }//end class


}