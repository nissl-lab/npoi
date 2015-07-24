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
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.IO;
    using System.Xml;
    using System.Xml.Serialization;

    /**
     * Holds details of built-in, default and user styles, which
     *  apply to tables / paragraphs / lists etc.
     * Text within one of those with custom stylings has the style
     *  information stored in the {@link XWPFRun}
     */
    public class XWPFStyles : POIXMLDocumentPart
    {
        private CT_Styles ctStyles;
        private List<XWPFStyle> listStyle = new List<XWPFStyle>();

        private XWPFLatentStyles latentStyles;
        private XWPFDefaultRunStyle defaultRunStyle;
        private XWPFDefaultParagraphStyle defaultParaStyle;

        /**
         * Construct XWPFStyles from a package part
         *
         * @param part the package part holding the data of the styles,
         * @param rel  the package relationship of type "http://schemas.Openxmlformats.org/officeDocument/2006/relationships/styles"
         */

        public XWPFStyles(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {
        }

        /**
         * Construct XWPFStyles from scratch for a new document.
         */
        public XWPFStyles()
        {
        }

        /**
         * Read document
         */

        internal override void OnDocumentRead()
        {
            StylesDocument stylesDoc;
            try
            {
                XmlDocument doc = ConvertStreamToXml(GetPackagePart().GetInputStream());
                stylesDoc = StylesDocument.Parse(doc,NamespaceManager);
                SetStyles(stylesDoc.Styles);
                latentStyles = new XWPFLatentStyles(ctStyles.latentStyles, this);

            }
            catch (XmlException e)
            {
                throw new POIXMLException("Unable to read styles", e);
            }
        }


        protected internal override void Commit()
        {
            if (ctStyles == null)
            {
                throw new InvalidOperationException("Unable to write out styles that were never read in!");
            }
            /*XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            xmlOptions.SaveSyntheticDocumentElement=(new QName(CTStyles.type.Name.NamespaceURI, "styles"));
            Dictionary<String,String> map = new Dictionary<String,String>();
            map.Put("http://schemas.Openxmlformats.org/officeDocument/2006/relationships", "r");
            map.Put("http://schemas.Openxmlformats.org/wordProcessingml/2006/main", "w");
            xmlOptions.SaveSuggestedPrefixes=(map);*/
            PackagePart part = GetPackagePart();
            using (Stream out1 = part.GetOutputStream())
            {
                StylesDocument doc = new StylesDocument(ctStyles);
                doc.Save(out1);
            }
        }

        protected void EnsureDocDefaults()
        {
            if (!ctStyles.IsSetDocDefaults())
            {
                ctStyles.AddNewDocDefaults();
            }

            CT_DocDefaults docDefaults = ctStyles.docDefaults;
            if (!docDefaults.IsSetPPrDefault())
                docDefaults.AddNewPPrDefault();
            if (!docDefaults.IsSetRPrDefault())
                docDefaults.AddNewRPrDefault();

            CT_PPrDefault pprd = docDefaults.pPrDefault;
            CT_RPrDefault rprd = docDefaults.rPrDefault;
            if (!pprd.IsSetPPr()) pprd.AddNewPPr();
            if (!rprd.IsSetRPr()) rprd.AddNewRPr();

            defaultRunStyle = new XWPFDefaultRunStyle(rprd.rPr);
            defaultParaStyle = new XWPFDefaultParagraphStyle(pprd.pPr);

        }

        /**
         * Sets the ctStyles
         * @param styles
         */
        public void SetStyles(CT_Styles styles)
        {
            ctStyles = styles;
            // Build up all the style objects
            foreach (CT_Style style in ctStyles.GetStyleList())
            {
                listStyle.Add(new XWPFStyle(style, this));
            }

            if (ctStyles.IsSetDocDefaults())
            {
                CT_DocDefaults docDefaults = ctStyles.docDefaults;
                if (docDefaults.IsSetRPrDefault() && docDefaults.rPrDefault.IsSetRPr())
                {
                    defaultRunStyle = new XWPFDefaultRunStyle(
                            docDefaults.rPrDefault.rPr);
                }
                if (docDefaults.IsSetPPrDefault() && docDefaults.pPrDefault.IsSetPPr())
                {
                    defaultParaStyle = new XWPFDefaultParagraphStyle(
                            docDefaults.pPrDefault.pPr);
                }
            }

        }

        /**
         * Checks whether style with styleID exist
         * @param styleID		styleID of the Style in the style-Document
         * @return				true if style exist, false if style not exist
         */
        public bool StyleExist(String styleID)
        {
            foreach (XWPFStyle style in listStyle)
            {
                if (style.StyleId.Equals(styleID))
                    return true;
            }
            return false;
        }
        /**
         * add a style to the document
         * @param style				
         * @throws IOException		 
         */
        public void AddStyle(XWPFStyle style)
        {
            listStyle.Add(style);
            ctStyles.AddNewStyle();
            int pos = (ctStyles.GetStyleList().Count) - 1;
            ctStyles.SetStyleArray(pos, style.GetCTStyle());
        }
        /**
         *get style by a styleID 
         * @param styleID	styleID of the searched style
         * @return style
         */
        public XWPFStyle GetStyle(String styleID)
        {
            foreach (XWPFStyle style in listStyle)
            {
                if (style.StyleId.Equals(styleID))
                    return style;
            }
            return null;
        }

        public int NumberOfStyles
        {
            get
            {
                return listStyle.Count;
            }
        }
        /**
         * Get the styles which are related to the parameter style and their relatives
         * this method can be used to copy all styles from one document to another document 
         * @param style
         * @return a list of all styles which were used by this method 
         */
        public List<XWPFStyle> GetUsedStyleList(XWPFStyle style)
        {
            List<XWPFStyle> usedStyleList = new List<XWPFStyle>();
            usedStyleList.Add(style);
            return GetUsedStyleList(style, usedStyleList);
        }

        /** 
         * Get the styles which are related to parameter style
         * @param style
         * @return all Styles of the parameterList
         */
        private List<XWPFStyle> GetUsedStyleList(XWPFStyle style, List<XWPFStyle> usedStyleList)
        {
            String basisStyleID = style.BasisStyleID;
            XWPFStyle basisStyle = GetStyle(basisStyleID);
            if ((basisStyle != null) && (!usedStyleList.Contains(basisStyle)))
            {
                usedStyleList.Add(basisStyle);
                GetUsedStyleList(basisStyle, usedStyleList);
            }
            String linkStyleID = style.LinkStyleID;
            XWPFStyle linkStyle = GetStyle(linkStyleID);
            if ((linkStyle != null) && (!usedStyleList.Contains(linkStyle)))
            {
                usedStyleList.Add(linkStyle);
                GetUsedStyleList(linkStyle, usedStyleList);
            }

            String nextStyleID = style.NextStyleID;
            XWPFStyle nextStyle = GetStyle(nextStyleID);
            if ((nextStyle != null) && (!usedStyleList.Contains(nextStyle)))
            {
                usedStyleList.Add(linkStyle);
                GetUsedStyleList(linkStyle, usedStyleList);
            }
            return usedStyleList;
        }

        protected CT_Language GetCTLanguage()
        {
            EnsureDocDefaults();

            CT_Language lang = null;
            if (defaultRunStyle.GetRPr().IsSetLang())
            {
                lang = defaultRunStyle.GetRPr().lang;
            }
            else
            {
                lang = defaultRunStyle.GetRPr().AddNewLang();
            }

            return lang;
        }

        /**
         * Sets the default spelling language on ctStyles DocDefaults parameter
         * @param strSpellingLanguage
         */
        public void SetSpellingLanguage(String strSpellingLanguage)
        {
            CT_Language lang = GetCTLanguage();
            lang.val = (/*setter*/strSpellingLanguage);
            lang.bidi = (/*setter*/strSpellingLanguage);
        }

        /**
         * Sets the default East Asia spelling language on ctStyles DocDefaults parameter
         * @param strEastAsia
         */
        public void SetEastAsia(String strEastAsia)
        {
            CT_Language lang = GetCTLanguage();
            lang.eastAsia = (/*setter*/strEastAsia);
        }

        /**
         * Sets the default font on ctStyles DocDefaults parameter
         * TODO Replace this with specific Setters for each type, possibly
         *  on XWPFDefaultRunStyle
         */
        public void SetDefaultFonts(CT_Fonts fonts)
        {
            EnsureDocDefaults();

            CT_RPr RunProps = defaultRunStyle.GetRPr();
            RunProps.rFonts = (/*setter*/fonts);
        }



        /**
         * Get the style with the same name
         * if this style is not existing, return null
         */
        public XWPFStyle GetStyleWithSameName(XWPFStyle style)
        {
            foreach (XWPFStyle ownStyle in listStyle)
            {
                if (ownStyle.HasSameName(style))
                {
                    return ownStyle;
                }
            }
            return null;

        }

        /**
         * Get the default style which applies text runs in the document
         */
        public XWPFDefaultRunStyle DefaultRunStyle
        {
            get
            {
                EnsureDocDefaults();
                return defaultRunStyle;
            }
        }
        /**
         * Get the default paragraph style which applies to the document
         */
        public XWPFDefaultParagraphStyle DefaultParagraphStyle
        {
            get
            {
                EnsureDocDefaults();
                return defaultParaStyle;
            }
        }

        /**
         * Get the definition of all the Latent Styles
         */
        public XWPFLatentStyles LatentStyles
        {
            get
            {
                return latentStyles;
            }
        }
    }

}