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

namespace NPOI.XSSF.UserModel.Extensions
{

    using NPOI.OpenXmlFormats.Spreadsheet;
    using System;
    using NPOI.XSSF.UserModel;
    using NPOI.XSSF.Model;
    using NPOI.SS.UserModel;
    /**
 * The enumeration value indicating the side being used for a cell border.
 */
    public enum BorderSide
    {
        TOP, RIGHT, BOTTOM, LEFT, DIAGONAL
    }

    /**
     * This element Contains border formatting information, specifying border defInition formats (left, right, top, bottom, diagonal)
     * for cells in the workbook.
     * Color is optional.
     */
    public class XSSFCellBorder
    {
        private ThemesTable _theme;
        private CT_Border border;

        /**
         * Creates a Cell Border from the supplied XML defInition
         */
        public XSSFCellBorder(CT_Border border, ThemesTable theme)
            : this(border)
        {

            this._theme = theme;
        }

        /**
         * Creates a Cell Border from the supplied XML defInition
         */
        public XSSFCellBorder(CT_Border border)
        {
            this.border = border;
        }

        /**
         * Creates a new, empty Cell Border.
         * You need to attach this to the Styles Table
         */
        public XSSFCellBorder()
        {
            border = new CT_Border();
        }

        /**
         * Records the Themes Table that is associated with
         *  the current font, used when looking up theme
         *  based colours and properties.
         */
        public void SetThemesTable(ThemesTable themes)
        {
            this._theme = themes;
        }


        /**
         * Returns the underlying XML bean.
         *
         * @return CT_Border
         */

        public CT_Border GetCTBorder()
        {
            return border;
        }

        /**
         * Get the type of border to use for the selected border
         *
         * @param side -  - where to apply the color defInition
         * @return borderstyle - the type of border to use. default value is NONE if border style is not Set.
         * @see BorderStyle
         */
        public BorderStyle GetBorderStyle(BorderSide side)
        {
            CT_BorderPr ctborder = GetBorder(side);
            ST_BorderStyle? border = ctborder == null ? ST_BorderStyle.none : ctborder.style;
            return (BorderStyle)border;
        }

        /**
         * Set the type of border to use for the selected border
         *
         * @param side  -  - where to apply the color defInition
         * @param style - border style
         * @see BorderStyle
         */
        public void SetBorderStyle(BorderSide side, BorderStyle style)
        {
            GetBorder(side, true).style = (ST_BorderStyle)Enum.GetValues(typeof(ST_BorderStyle)).GetValue((int)style + 1);
        }

        /**
         * Get the color to use for the selected border
         *
         * @param side - where to apply the color defInition
         * @return color - color to use as XSSFColor. null if color is not set
         */
        public XSSFColor GetBorderColor(BorderSide side)
        {
            CT_BorderPr borderPr = GetBorder(side);

            if (borderPr != null && borderPr.IsSetColor())
            {
                XSSFColor clr = new XSSFColor(borderPr.color);
                if (_theme != null)
                {
                    _theme.InheritFromThemeAsRequired(clr);
                }
                return clr;
            }
            else
            {
                // No border set
                return null;
            }
        }

        /**
         * Set the color to use for the selected border
         *
         * @param side  - where to apply the color defInition
         * @param color - the color to use
         */
        public void SetBorderColor(BorderSide side, XSSFColor color)
        {
            CT_BorderPr borderPr = GetBorder(side, true);
            if (color == null) borderPr.UnsetColor();
            else
                borderPr.color = color.GetCTColor();
        }

        private CT_BorderPr GetBorder(BorderSide side)
        {
            return GetBorder(side, false);
        }


        private CT_BorderPr GetBorder(BorderSide side, bool ensure)
        {
            CT_BorderPr borderPr;
            switch (side)
            {
                case BorderSide.TOP:
                    borderPr = border.top;
                    if (ensure && borderPr == null) borderPr = border.AddNewTop();
                    break;
                case BorderSide.RIGHT:
                    borderPr = border.right;
                    if (ensure && borderPr == null) borderPr = border.AddNewRight();
                    break;
                case BorderSide.BOTTOM:
                    borderPr = border.bottom;
                    if (ensure && borderPr == null) borderPr = border.AddNewBottom();
                    break;
                case BorderSide.LEFT:
                    borderPr = border.left;
                    if (ensure && borderPr == null) borderPr = border.AddNewLeft();
                    break;
                case BorderSide.DIAGONAL:
                    borderPr = border.diagonal;
                    if (ensure && borderPr == null) borderPr = border.AddNewDiagonal();
                    break;
                default:
                    throw new ArgumentException("No suitable side specified for the border");
            }
            return borderPr;
        }


        public override int GetHashCode()
        {
            return border.ToString().GetHashCode();
        }

        public override bool Equals(Object o)
        {
            if (!(o is XSSFCellBorder)) return false;

            //TODO: change the compare logic
            XSSFCellBorder cf = (XSSFCellBorder)o;
            return border.ToString().Equals(cf.GetCTBorder().ToString());
        }
    }
}


