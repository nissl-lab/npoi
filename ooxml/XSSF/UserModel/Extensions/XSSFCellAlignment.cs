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
using NPOI.SS.UserModel;
using System;
namespace NPOI.XSSF.UserModel.Extensions
{

    /**
     * Cell Settings avaiable in the Format/Alignment tab
     */
    public class XSSFCellAlignment
    {

        private CT_CellAlignment cellAlignement;

        /**
         * Creates a Cell Alignment from the supplied XML defInition
         *
         * @param cellAlignment
         */
        public XSSFCellAlignment(CT_CellAlignment cellAlignment)
        {
            this.cellAlignement = cellAlignment;
        }

        /**
         * Get the type of vertical alignment for the cell
         *
         * @return the type of aligment
         * @see VerticalAlignment
         */
        public VerticalAlignment GetVertical()
        {
            ST_VerticalAlignment align = cellAlignement.vertical;

            return (VerticalAlignment)align;
        }

        /**
         * Set the type of vertical alignment for the cell
         *
         * @param align - the type of alignment
         * @see VerticalAlignment
         */
        public void SetVertical(VerticalAlignment align)
        {
            cellAlignement.vertical = (ST_VerticalAlignment)align;
            cellAlignement.verticalSpecified = true;
        }

        /**
         * Get the type of horizontal alignment for the cell
         *
         * @return the type of aligment
         * @see HorizontalAlignment
         */
        public HorizontalAlignment GetHorizontal()
        {
            ST_HorizontalAlignment align = cellAlignement.horizontal;
            

            return (HorizontalAlignment)align;
        }

        /**
         * Set the type of horizontal alignment for the cell
         *
         * @param align - the type of alignment
         * @see HorizontalAlignment
         */
        public void SetHorizontal(HorizontalAlignment align)
        {
            cellAlignement.horizontal = ((ST_HorizontalAlignment)align);
            cellAlignement.horizontalSpecified = true;
        }

        /**
         * Get the number of spaces to indent the text in the cell
         *
         * @return indent - number of spaces
         */
        public long GetIndent()
        {
            return cellAlignement.indent;
        }

        /**
         * Set the number of spaces to indent the text in the cell
         *
         * @param indent - number of spaces
         */
        public void SetIndent(long indent)
        {
            cellAlignement.indent = (indent);
            cellAlignement.indentSpecified = true;
        }

        /**
         * Get the degree of rotation for the text in the cell
         * <p/>
         * Expressed in degrees. Values range from 0 to 180. The first letter of
         * the text is considered the center-point of the arc.
         * <br/>
         * For 0 - 90, the value represents degrees above horizon. For 91-180 the degrees below the
         * horizon is calculated as:
         * <br/>
         * <code>[degrees below horizon] = 90 - textRotation.</code>
         * </p>
         *
         * @return rotation degrees (between 0 and 180 degrees)
         */
        public long GetTextRotation()
        {
            return cellAlignement.textRotation;
        }

        /**
         * Set the degree of rotation for the text in the cell
         * <p/>
         * Expressed in degrees. Values range from 0 to 180. The first letter of
         * the text is considered the center-point of the arc.
         * <br/>
         * For 0 - 90, the value represents degrees above horizon. For 91-180 the degrees below the
         * horizon is calculated as:
         * <br/>
         * <code>[degrees below horizon] = 90 - textRotation.</code>
         * </p>
         *
         * @param rotation - the rotation degrees (between 0 and 180 degrees)
         */
        public void SetTextRotation(long rotation)
        {
            cellAlignement.textRotation = (rotation);
            cellAlignement.textRotationSpecified = true;
        }

        /**
         * Whether the text should be wrapped
         *
         * @return a bool value indicating if the text in a cell should be line-wrapped within the cell.
         */
        public bool GetWrapText()
        {
            return cellAlignement.wrapText;
        }

        /**
         * Set whether the text should be wrapped
         *
         * @param wrapped a bool value indicating if the text in a cell should be line-wrapped within the cell.
         */
        public void SetWrapText(bool wrapped)
        {
            cellAlignement.wrapText = (wrapped);
            cellAlignement.wrapTextSpecified = true;
        }

        /**
         * Access to low-level data
         */

        public CT_CellAlignment GetCTCellAlignment()
        {
            return cellAlignement;
        }
    }

}

