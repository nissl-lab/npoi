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
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml;
using System;
using System.IO;
using System.Xml;

namespace NPOI.XWPF.UserModel
{
    /// <summary>
    ///  A shared style sheet in a .docx document
    /// </summary>
    public class XWPFTheme : POIXMLDocumentPart
    {
        private CT_OfficeStyleSheet theme;

        // Theme element index constants - same semantics as spreadsheet ThemesTable
        public const int THEME_LT1 = 0;
        public const int THEME_DK1 = 1;
        public const int THEME_LT2 = 2;
        public const int THEME_DK2 = 3;
        public const int THEME_ACCENT1 = 4;
        public const int THEME_ACCENT2 = 5;
        public const int THEME_ACCENT3 = 6;
        public const int THEME_ACCENT4 = 7;
        public const int THEME_ACCENT5 = 8;
        public const int THEME_ACCENT6 = 9;
        public const int THEME_HLINK = 10;
        public const int THEME_FOLHLINK = 11;

        public XWPFTheme()
            : base()
        {
            theme = new CT_OfficeStyleSheet();
        }

        public XWPFTheme(PackagePart part)
            : base(part)
        {
            using Stream inst = part.GetInputStream();
            XmlDocument xmldoc = ConvertStreamToXml(inst);
            try
            {
                theme = CT_OfficeStyleSheet.Parse(xmldoc.DocumentElement, NamespaceManager);
            }
            catch(XmlException e)
            {
                throw new IOException(e.Message, e);
            }
        }

        internal XWPFTheme(CT_OfficeStyleSheet theme)
        {
            this.theme = theme;
        }
        /// <summary>
        /// name of this theme, e.g. "Office Theme"
        /// </summary>
        public String Name
        {
            get
            {
                return theme.name;
            }
            set {
                this.theme.name=value;
            }
        }
        public CT_Color GetCTColor(String name)
        {
            CT_ColorScheme scheme = theme.themeElements?.clrScheme;
            return GetMapColor(name, scheme);
        }


        private static CT_Color GetMapColor(String mapName, CT_ColorScheme scheme)
        {
            if(mapName == null || scheme == null)
            {
                return null;
            }
            switch(mapName)
            {
                case "accent1":
                    return scheme.accent1;
                case "accent2":
                    return scheme.accent2;
                case "accent3":
                    return scheme.accent3;
                case "accent4":
                    return scheme.accent4;
                case "accent5":
                    return scheme.accent5;
                case "accent6":
                    return scheme.accent6;
                case "dk1":
                    return scheme.dk1;
                case "dk2":
                    return scheme.dk2;
                case "folHlink":
                    return scheme.folHlink;
                case "hlink":
                    return scheme.hlink;
                case "lt1":
                    return scheme.lt1;
                case "lt2":
                    return scheme.lt2;
                default:
                    return null;
            }
        }


        /// <summary>
        /// CT_OfficeStyleSheet（drawingml theme）
        /// </summary>
        public CT_OfficeStyleSheet GetCTTheme()
        {
            return theme;
        }

        /// <summary>
        /// Typically the major font is used for heading areas of a document.
        /// </summary>
        public String MajorFont
        {
            get
            {
                return theme?.themeElements?.fontScheme?.majorFont?.latin?.typeface;
            }

        }
        /// <summary>
        /// Typically the monor font is used for normal text or paragraph areas.
        /// </summary>
        public String MinorFont
        {
            get
            {
                return theme?.themeElements?.fontScheme?.minorFont?.latin?.typeface;
            }
        }

        protected internal override void Commit()
        {
            if(theme == null)
            {
                throw new IOException("Unable to write out theme that was never read in!");
            }
            PackagePart part = GetPackagePart();
            using Stream out1 = part.GetOutputStream();
            using StreamWriter sw = new StreamWriter(out1, System.Text.Encoding.UTF8, 1024, true);
            theme.Write(sw);
        }
        public void SetTheme(CT_OfficeStyleSheet theme)
        {
            this.theme = theme;
        }
    }
}