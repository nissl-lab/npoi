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
namespace NPOI.XSSF.Model
{

    using System.Xml;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.XSSF.UserModel;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.OpenXmlFormats.Dml;
    using System;
    using System.Collections.Generic;
    using System.IO;

    /**
     * Class that represents theme of XLSX document. The theme includes specific
     * colors and fonts.
     */
    public class ThemesTable : POIXMLDocumentPart
    {
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

        private ThemeDocument theme;
        /**
     * Create a new, empty ThemesTable
     */
        public ThemesTable() :
            base()
        {
            theme = new ThemeDocument();
            theme.AddNewTheme().AddNewThemeElements();
        }
        /**
         * Construct a ThemesTable.
         * @param part A PackagePart.
         * @param rel A PackageRelationship.
         */
        public ThemesTable(PackagePart part)
            : base(part)
        {

            XmlDocument xmldoc = ConvertStreamToXml(part.GetInputStream());

            try
            {
                theme = ThemeDocument.Parse(xmldoc, NamespaceManager);
            }
            catch (XmlException e)
            {
                throw new IOException(e.Message, e);
            }
        }

        [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        public ThemesTable(PackagePart part, PackageRelationship rel)
             : this(part)
        {

        }
        /**
         * Construct a ThemesTable from an existing ThemeDocument.
         * @param theme A ThemeDocument.
         */
        internal ThemesTable(ThemeDocument theme)
        {
            this.theme = theme;
        }
        /**
         * Convert a theme "index" into a color.
         * @param idx A theme "index"
         * @return The mapped XSSFColor, or null if not mapped.
         */
        public XSSFColor GetThemeColor(int idx)
        {
            // Theme color references are NOT positional indices into the color scheme,
            // i.e. these keys are NOT the same as the order in which theme colors appear
            // in theme1.xml. They are keys to a mapped color.
            CT_ColorScheme colorScheme = theme.GetTheme().themeElements.clrScheme;
            NPOI.OpenXmlFormats.Dml.CT_Color ctColor = null;

            switch (idx)
            {
                case THEME_LT1: ctColor = colorScheme.lt1; break;
                case THEME_DK1: ctColor = colorScheme.dk1; break;
                case THEME_LT2: ctColor = colorScheme.lt2; break;
                case THEME_DK2: ctColor = colorScheme.dk2; break;
                case THEME_ACCENT1: ctColor = colorScheme.accent1; break;
                case THEME_ACCENT2: ctColor = colorScheme.accent2; break;
                case THEME_ACCENT3: ctColor = colorScheme.accent3; break;
                case THEME_ACCENT4: ctColor = colorScheme.accent4; break;
                case THEME_ACCENT5: ctColor = colorScheme.accent5; break;
                case THEME_ACCENT6: ctColor = colorScheme.accent6; break;
                case THEME_HLINK: ctColor = colorScheme.hlink; break;
                case THEME_FOLHLINK: ctColor = colorScheme.folHlink; break;
                default: return null;
            }

            byte[] rgb = null;
            if (ctColor.IsSetSrgbClr())
            {
                // Color is a regular one
                rgb = ctColor.srgbClr.val;
            }
            else if (ctColor.IsSetSysClr())
            {
                // Color is a tint of white or black
                rgb = ctColor.sysClr.lastClr;
            }
            else
            {
                return null;
            }
            return new XSSFColor(rgb);

        }

        /**
         * If the colour is based on a theme, then inherit 
         *  information (currently just colours) from it as
         *  required.
         */
        public void InheritFromThemeAsRequired(XSSFColor color)
        {
            if (color == null)
            {
                // Nothing for us to do
                return;
            }
            if (!color.GetCTColor().themeSpecified)
            {
                // No theme Set, nothing to do
                return;
            }

            // Get the theme colour
            XSSFColor themeColor = GetThemeColor(color.Theme);
            // Set the raw colour, not the adjusted one
            // Do a raw Set, no adjusting at the XSSFColor layer either
            color.GetCTColor().SetRgb(themeColor.GetCTColor().GetRgb());

            // All done
        }

        /**
         * Write this table out as XML.
         * 
         * @param out The stream to write to.
         * @throws IOException if an error occurs while writing.
         */
        public void writeTo(Stream out1)
        {
            //XmlOptions options = new XmlOptions(DEFAULT_XML_OPTIONS);

            theme.Save(out1);
        }

        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            writeTo(out1);
            out1.Close();
        }
    }

    public class ThemeElement
    {

        private static SortedDictionary<int, ThemeElement> values = new SortedDictionary<int, ThemeElement>();
        public static ThemeElement LT1 = new ThemeElement(0, "Lt1");
        public static ThemeElement DK1 = new ThemeElement(1, "Dk1");
        public static ThemeElement LT2 = new ThemeElement(2, "Lt2");
        public static ThemeElement DK2 = new ThemeElement(3, "Dk2");
        public static ThemeElement ACCENT1 = new ThemeElement(4, "Accent1");
        public static ThemeElement ACCENT2 = new ThemeElement(5, "Accent2");
        public static ThemeElement ACCENT3 = new ThemeElement(6, "Accent3");
        public static ThemeElement ACCENT4 = new ThemeElement(7, "Accent4");
        public static ThemeElement ACCENT5 = new ThemeElement(8, "Accent5");
        public static ThemeElement ACCENT6 = new ThemeElement(9, "Accent6");
        public static ThemeElement HLINK = new ThemeElement(10, "Hlink");
        public static ThemeElement FOLHLINK = new ThemeElement(11, "FolHlink");
        public static ThemeElement UNKNOWN = new ThemeElement(-1,null);

        public static ThemeElement ById(int idx)
        {
            if (idx >= values.Count || idx < 0) return UNKNOWN;
            return values[idx];
        }
        private ThemeElement(int idx, String name)
        {
            this.idx = idx; this.name = name;
            values.Add(idx, this);
        }
        public int idx;
        public String name;
   }
}




