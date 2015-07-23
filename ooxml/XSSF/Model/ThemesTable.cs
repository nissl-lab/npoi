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
        private ThemeDocument theme;
        /**
         * Construct a ThemesTable.
         * @param part A PackagePart.
         * @param rel A PackageRelationship.
         */
        internal ThemesTable(PackagePart part, PackageRelationship rel)
            : base(part, rel)
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
                case 0: ctColor = colorScheme.lt1; break;
                case 1: ctColor = colorScheme.dk1; break;
                case 2: ctColor = colorScheme.lt2; break;
                case 3: ctColor = colorScheme.dk2; break;
                case 4: ctColor = colorScheme.accent1; break;
                case 5: ctColor = colorScheme.accent2; break;
                case 6: ctColor = colorScheme.accent3; break;
                case 7: ctColor = colorScheme.accent4; break;
                case 8: ctColor = colorScheme.accent5; break;
                case 9: ctColor = colorScheme.accent6; break;
                case 10: ctColor = colorScheme.hlink; break;
                case 11: ctColor = colorScheme.folHlink; break;
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
    }

}




