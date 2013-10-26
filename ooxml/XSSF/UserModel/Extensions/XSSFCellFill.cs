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
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.OpenXmlFormats;
namespace NPOI.XSSF.UserModel.Extensions
{

    /**
     * This element specifies fill formatting.
     * A cell fill consists of a background color, foreground color, and pattern to be applied across the cell.
     */
    public class XSSFCellFill
    {

        private CT_Fill _fill;

        /**
         * Creates a CellFill from the supplied parts
         *
         * @param fill - fill
         */
        public XSSFCellFill(CT_Fill fill)
        {
            _fill = fill;
        }

        /**
         * Creates an empty CellFill
         */
        public XSSFCellFill()
        {
            _fill = new CT_Fill();
        }

        /**
         * Get the background fill color.
         *
         * @return fill color, null if color is not set
         */
        public XSSFColor GetFillBackgroundColor()
        {
            CT_PatternFill ptrn = _fill.GetPatternFill();
            if (ptrn == null) return null;

            CT_Color CT_Color = ptrn.bgColor;
            return CT_Color == null ? null : new XSSFColor(CT_Color);
        }

        /**
         * Set the background fill color represented as a indexed color value.
         *
         * @param index
         */
        public void SetFillBackgroundColor(int index)
        {
            CT_PatternFill ptrn = EnsureCTPatternFill();
            CT_Color ctColor = ptrn.IsSetBgColor() ? ptrn.bgColor : ptrn.AddNewBgColor();
            ctColor.indexed = (uint)index;
            ctColor.indexedSpecified = true;

        }

        /**
         * Set the background fill color represented as a {@link XSSFColor} value.
         *
         * @param color
         */
        public void SetFillBackgroundColor(XSSFColor color)
        {
            CT_PatternFill ptrn = EnsureCTPatternFill();
            ptrn.bgColor = (color.GetCTColor());
        }

        /**
         * Get the foreground fill color.
         *
         * @return XSSFColor - foreground color. null if color is not set
         */
        public XSSFColor GetFillForegroundColor()
        {
            CT_PatternFill ptrn = _fill.GetPatternFill();
            if (ptrn == null) return null;

            CT_Color ctColor = ptrn.fgColor;
            return ctColor == null ? null : new XSSFColor(ctColor);
        }

        /**
         * Set the foreground fill color as a indexed color value
         *
         * @param index - the color to use
         */
        public void SetFillForegroundColor(int index)
        {
            CT_PatternFill ptrn = EnsureCTPatternFill();
            CT_Color CT_Color = ptrn.IsSetFgColor() ? ptrn.fgColor : ptrn.AddNewFgColor();
            CT_Color.indexed = (uint)index;
        }

        /**
         * Set the foreground fill color represented as a {@link XSSFColor} value.
         *
         * @param color - the color to use
         */
        public void SetFillForegroundColor(XSSFColor color)
        {
            CT_PatternFill ptrn = EnsureCTPatternFill();
            ptrn.fgColor = color.GetCTColor();
        }

        /**
         * get the fill pattern
         *
         * @return fill pattern type. null if fill pattern is not set
         */
        public ST_PatternType GetPatternType()
        {
            CT_PatternFill ptrn = _fill.GetPatternFill();
            return ptrn == null ? ST_PatternType.none : (ST_PatternType)ptrn.patternType;
        }

        /**
         * set the fill pattern
         *
         * @param patternType fill pattern to use
         */
        public void SetPatternType(ST_PatternType patternType)
        {
            CT_PatternFill ptrn = EnsureCTPatternFill();
            ptrn.patternType = patternType;
        }

        private CT_PatternFill EnsureCTPatternFill()
        {
            CT_PatternFill patternFill = _fill.GetPatternFill();
            if (patternFill == null)
            {
                patternFill = _fill.AddNewPatternFill();
            }
            return patternFill;
        }

        /**
         * Returns the underlying XML bean.
         *
         * @return CT_Fill
         */

        internal CT_Fill GetCTFill()
        {
            return _fill;
        }


        public override int GetHashCode()
        {
            return _fill.ToString().GetHashCode();
        }

        public override bool Equals(object o)
        {
            if (!(o is XSSFCellFill)) return false;

            XSSFCellFill cf = (XSSFCellFill)o;
            return _fill.ToString().Equals(cf.GetCTFill().ToString());
        }
    }
}