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
    using NPOI.OpenXml4Net.OPC;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.IO;
    using System.Xml.Serialization;
    using System.Xml;
    using NPOI.Util;
    using System.Xml.XPath;

    /**
     * Sketch of XWPF header class
     */
    public class XWPFHeader : XWPFHeaderFooter
    {
        public XWPFHeader():base()
        {
            headerFooter = new CT_Hdr();
            ReadHdrFtr();
        }

        public XWPFHeader(POIXMLDocumentPart parent, PackagePart part, PackageRelationship rel)
            : base(parent, part, rel)
        {
            
        }

        public XWPFHeader(XWPFDocument doc, CT_HdrFtr hdrFtr)
            : base(doc, hdrFtr)
        {
            /*
            XmlCursor cursor = headerFooter.NewCursor();
            cursor.SelectPath("./*");
            while (cursor.ToNextSelection()) {
                XmlObject o = cursor.Object;
                if (o is CTP) {
                    XWPFParagraph p = new XWPFParagraph((CTP) o, this);
                    paragraphs.Add(p);
                }
                if (o is CTTbl) {
                    XWPFTable t = new XWPFTable((CTTbl) o, this);
                    tables.Add(t);
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

        /// <summary>
        /// Save and commit footer
        /// </summary>
        protected internal override void Commit()
        {
            /*XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            xmlOptions.SaveSyntheticDocumentElement=(new QName(CTNumbering.type.Name.NamespaceURI, "hdr"));
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
            xmlOptions.SaveSuggestedPrefixes=(/map);*/
            PackagePart part = GetPackagePart();
            using (Stream out1 = part.GetOutputStream())
            {
                HdrDocument doc = new HdrDocument((CT_Hdr)headerFooter);
                doc.Save(out1);
            }
        }


        /// <summary>
        /// Read the document
        /// </summary>
        internal override void OnDocumentRead()
        {
            base.OnDocumentRead();
            HdrDocument hdrDocument = null;
            try
            {
                XmlDocument xmldoc = DocumentHelper.LoadDocument(GetPackagePart().GetInputStream());
                hdrDocument = HdrDocument.Parse(xmldoc, NamespaceManager);
                headerFooter = hdrDocument.Hdr;
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
                    if (o is CT_SdtBlock)
                    {
                        XWPFSDT c = new XWPFSDT((CT_SdtBlock)o, this);
                        bodyElements.Add(c);
                    }
                }
            }
            catch (Exception e)
            {
                throw new POIXMLException(e);
            }
        }
        /// <summary>
        /// Get the PartType of the body
        /// </summary>
        public override BodyType PartType
        {
            get
            {
                return BodyType.HEADER;
            }
        }
    }
}