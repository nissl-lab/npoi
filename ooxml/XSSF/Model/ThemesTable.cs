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
     * 
     * @author Petr Udalau(Petr.Udalau at exigenservices.com) - theme colors
     */
    public class ThemesTable : POIXMLDocumentPart
    {
        private ThemeDocument theme;

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
                throw new IOException(e.Message);
            }
        }

        internal ThemesTable(ThemeDocument theme)
        {
            this.theme = theme;
        }

        public XSSFColor GetThemeColor(int idx)
        {
            CT_ColorScheme colorScheme = theme.GetTheme().themeElements.clrScheme;
            NPOI.OpenXmlFormats.Dml.CT_Color ctColor = null;
            
            int cnt = 0;
            List<NPOI.OpenXmlFormats.Dml.CT_Color> colors = new List<NPOI.OpenXmlFormats.Dml.CT_Color>();
            colors.Add(colorScheme.dk1);
            colors.Add(colorScheme.lt1);
            colors.Add(colorScheme.dk2);
            colors.Add(colorScheme.lt2);
            colors.Add(colorScheme.accent1);
            colors.Add(colorScheme.accent2);
            colors.Add(colorScheme.accent3);
            colors.Add(colorScheme.accent4);
            colors.Add(colorScheme.accent5);
            colors.Add(colorScheme.accent6);
            colors.Add(colorScheme.hlink);
            colors.Add(colorScheme.folHlink);

            //iterate ctcolors in colorschema
            foreach (NPOI.OpenXmlFormats.Dml.CT_Color color in colors)
            {
                    if (cnt == idx)
                    {
                        ctColor = color;

                        byte[] rgb = null;
                        if (ctColor.srgbClr != null)
                        {
                            // Colour is a regular one 
                            rgb = ctColor.srgbClr.val;
                        }
                        else if (ctColor.sysClr != null)
                        {
                            // Colour is a tint of white or black
                            rgb = ctColor.sysClr.lastClr;
                        }

                        return new XSSFColor(rgb);
                    }
                    cnt++;
            }
            return null;
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
            XSSFColor themeColor = GetThemeColor(color.GetTheme());
            // Set the raw colour, not the adjusted one
            // Do a raw Set, no adjusting at the XSSFColor layer either
            color.GetCTColor().SetRgb(themeColor.GetCTColor().GetRgb());

            // All done
        }
    }

}




