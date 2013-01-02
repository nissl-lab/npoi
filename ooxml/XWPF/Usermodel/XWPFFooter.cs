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
    using NPOI.OpenXml4Net.OPC;
    using System.IO;
    using System.Xml;
    using System.Xml.Serialization;

    /**
     * Sketch of XWPF footer class
     */
    public class XWPFFooter : XWPFHeaderFooter
    {
        public XWPFFooter()
            //: base()
        {
            headerFooter = new CT_Ftr();
            ReadHdrFtr();
        }

        public XWPFFooter(XWPFDocument doc, CT_HdrFtr hdrFtr)
            : base(doc, hdrFtr)
        {
            /*
            XmlCursor cursor = headerFooter.NewCursor();
            cursor.SelectPath("./*");
            while (cursor.ToNextSelection()) {
                XmlObject o = cursor.Object;
                if (o is CT_P) {
                    XWPFParagraph p = new XWPFParagraph((CT_P)o, this);
                    paragraphs.Add(p);
                    bodyElements.Add(p);
                }
                if (o is CT_Tbl) {
                    XWPFTable t = new XWPFTable((CT_Tbl)o, this);
                    tables.Add(t);
                    bodyElements.Add(t);
                }
            }
            cursor.Dispose();*/
            foreach (object o in hdrFtr.Items)
            {
                if (o is CT_P)
                {
                    XWPFParagraph p = new XWPFParagraph((CT_P)o, this);
                    paragraphs.Add(p);
                }
                if (o is CT_Tbl)
                {
                    XWPFTable t = new XWPFTable((CT_Tbl)o, this);
                    tables.Add(t);
                }
            }
        }

        public XWPFFooter(POIXMLDocumentPart parent, PackagePart part, PackageRelationship rel)
            : base(parent, part, rel)
        {
        }

        /**
         * save and Commit footer
         */

        protected override void Commit()
        {
            /*XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            xmlOptions.SaveSyntheticDocumentElement=(new QName(CTNumbering.type.Name.NamespaceURI, "ftr"));
            Dictionary<String,String> map = new Dictionary<String, String>();
            map.Put("http://schemas.Openxmlformats.org/markup-compatibility/2006", "ve");
            map.Put("urn:schemas-microsoft-com:office:office", "o");
            map.Put("http://schemas.Openxmlformats.org/officeDocument/2006/relationships", "r");
            map.Put("http://schemas.Openxmlformats.org/officeDocument/2006/math", "m");
            map.Put("urn:schemas-microsoft-com:vml", "v");
            map.Put("http://schemas.Openxmlformats.org/drawingml/2006/wordProcessingDrawing", "wp");
            map.Put("urn:schemas-microsoft-com:office:word", "w10");
            map.Put("http://schemas.Openxmlformats.org/wordProcessingml/2006/main", "w");
            map.Put("http://schemas.microsoft.com/office/word/2006/wordml", "wne");
            xmlOptions.SaveSuggestedPrefixes=(map);*/
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            FtrDocument doc = new FtrDocument((CT_Ftr)headerFooter);
            XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
                new XmlQualifiedName("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
                new XmlQualifiedName("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
                new XmlQualifiedName("m", "http://schemas.openxmlformats.org/officeDocument/2006/math"),
                new XmlQualifiedName("v", "urn:schemas-microsoft-com:vml"),
                new XmlQualifiedName("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
                new XmlQualifiedName("w10", "urn:schemas-microsoft-com:office:word"),
                new XmlQualifiedName("wne", "http://schemas.microsoft.com/office/word/2006/wordml"),
                 new XmlQualifiedName("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
             });
            doc.Save(out1, namespaces);
            out1.Close();
        }


        internal override void OnDocumentRead()
        {
            base.OnDocumentRead();
            FtrDocument ftrDocument = null;
            Stream is1;
            try {
                is1 = GetPackagePart().GetInputStream();
                ftrDocument = FtrDocument.Parse(is1);
                headerFooter = ftrDocument.Ftr;
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
                while (cursor.ToNextSelection()) {
                    XmlObject o = cursor.Object;
                    if (o is CTP) {
                        XWPFParagraph p = new XWPFParagraph((CTP)o, this);
                        paragraphs.Add(p);
                        bodyElements.Add(p);
                    }
                    if (o is CTTbl) {
                        XWPFTable t = new XWPFTable((CTTbl)o, this);
                        tables.Add(t);
                        bodyElements.Add(t);
                    }
                }
                cursor.Dispose();*/
            } catch (Exception e) {
                throw new POIXMLException(e);
            }
        }

        /**
         * Get the PartType of the body
         * @see NPOI.XWPF.UserModel.IBody#getPartType()
         */
        public override BodyType GetPartType()
        {
            return BodyType.FOOTER;
        }
    }

}