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
    /**
     * @author Yegor Kozlov
     */
    public class XWPFRelation : POIXMLRelation
    {

        /**
         * A map to lookup POIXMLRelation by its relation type
         */
        protected static Dictionary<String, XWPFRelation> _table = new Dictionary<String, XWPFRelation>();


        public static XWPFRelation DOCUMENT = new XWPFRelation(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "/word/document.xml",
                null
        );
        public static XWPFRelation TEMPLATE = new XWPFRelation(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
          "/word/document.xml",
              null
        );
        public static XWPFRelation MACRO_DOCUMENT = new XWPFRelation(
                "application/vnd.ms-word.document.macroEnabled.main+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "/word/document.xml",
                null
        );
        public static XWPFRelation MACRO_TEMPLATE_DOCUMENT = new XWPFRelation(
                "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "/word/document.xml",
            null
        );
        public static XWPFRelation GLOSSARY_DOCUMENT = new XWPFRelation(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument",
            "/word/glossary/document.xml",
            null
        );
        public static XWPFRelation NUMBERING = new XWPFRelation(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
            "/word/numbering.xml",
                typeof(XWPFNumbering)
        );
        public static XWPFRelation FONT_TABLE = new XWPFRelation(
               "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
            "/word/fontTable.xml",
                null
        );
        public static XWPFRelation SETTINGS = new XWPFRelation(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
            "/word/settings.xml",
                typeof(XWPFSettings)
        );
        public static XWPFRelation STYLES = new XWPFRelation(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            "/word/styles.xml",
                typeof(XWPFStyles)
        );
        public static XWPFRelation WEB_SETTINGS = new XWPFRelation(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
            "/word/webSettings.xml",
                null
        );
        public static XWPFRelation HEADER = new XWPFRelation(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
            "/word/header#.xml",
                typeof(XWPFHeader)
        );
        public static XWPFRelation FOOTER = new XWPFRelation(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
            "/word/footer#.xml",
                typeof(XWPFFooter)
        );

        public static XWPFRelation THEME = new XWPFRelation(
            "application/vnd.openxmlformats-officedocument.theme+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
            "/word/theme/theme#.xml",
            null
        );
        public static XWPFRelation HYPERLINK = new XWPFRelation(
                null,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                null,
                null
        );
        public static XWPFRelation COMMENT = new XWPFRelation(
                null,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                null,
                null
        );
        public static XWPFRelation FOOTNOTE = new XWPFRelation(
               "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
           "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
           "/word/footnotes.xml",
               typeof(XWPFFootnotes)
        );
        public static XWPFRelation ENDNOTE = new XWPFRelation(
                null,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
                null,
                null
        );
        /**
         * Supported image formats
         */
        public static XWPFRelation IMAGE_EMF = new XWPFRelation(
              "image/x-emf",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          "/word/media/image#.emf",
              typeof(XWPFPictureData)
        );
        public static XWPFRelation IMAGE_WMF = new XWPFRelation(
              "image/x-wmf",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          "/word/media/image#.wmf",
              typeof(XWPFPictureData)
        );
        public static XWPFRelation IMAGE_PICT = new XWPFRelation(
              "image/pict",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          "/word/media/image#.pict",
              typeof(XWPFPictureData)
        );
        public static XWPFRelation IMAGE_JPEG = new XWPFRelation(
              "image/jpeg",
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              "/word/media/image#.jpeg",
              typeof(XWPFPictureData)
        );
        public static XWPFRelation IMAGE_PNG = new XWPFRelation(
              "image/png",
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              "/word/media/image#.png",
              typeof(XWPFPictureData)
        );
        public static XWPFRelation IMAGE_DIB = new XWPFRelation(
              "image/dib",
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              "/word/media/image#.dib",
              typeof(XWPFPictureData)
        );
        public static XWPFRelation IMAGE_GIF = new XWPFRelation(
              "image/gif",
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              "/word/media/image#.gif",
              typeof(XWPFPictureData)
        );

        public static XWPFRelation IMAGE_TIFF = new XWPFRelation(
            "image/tiff",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            "/word/media/image#.tiff",
            typeof(XWPFPictureData)
        );
        public static XWPFRelation IMAGE_EPS = new XWPFRelation(
                "image/x-eps",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                "/word/media/image#.eps",
                typeof(XWPFPictureData)
        );
        public static XWPFRelation IMAGE_BMP = new XWPFRelation(
                "image/x-ms-bmp",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                "/word/media/image#.bmp",
                typeof(XWPFPictureData)
        );
        public static XWPFRelation IMAGE_WPG = new XWPFRelation(
                "image/x-wpg",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                "/word/media/image#.wpg",
                typeof(XWPFPictureData)
        );

        public static XWPFRelation IMAGES = new XWPFRelation(
              null,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              null,
              null
        );


        private XWPFRelation(String type, String rel, String defaultName, Type cls)
            : base(type, rel, defaultName, cls)
        {
            ;

            if (cls != null && !_table.ContainsKey(rel)) _table.Add(rel, this);
        }

        /**
         * Get POIXMLRelation by relation type
         *
         * @param rel relation type, for example,
         *            <code>http://schemas.openxmlformats.org/officeDocument/2006/relationships/image</code>
         * @return registered POIXMLRelation or null if not found
         */
        public static XWPFRelation GetInstance(String rel)
        {
            if (_table.ContainsKey(rel))
                return _table[(rel)];
            return null;
        }

    }

}