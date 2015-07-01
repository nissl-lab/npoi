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

namespace NPOI.SS.UserModel
{
    using System;

    public interface ICellStyle
    {
        /// <summary>
        /// the Cell should be auto-sized to shrink to fit if the text is too long
        /// </summary>
        bool ShrinkToFit { get; set; }
        /**
         * get the index within the Workbook (sequence within the collection of ExtnededFormat objects)
         * @return unique index number of the underlying record this style represents (probably you don't care
         *  unless you're comparing which one is which)
         */

        short Index { get; }

        /**
         * get the index of the format
         * @see DataFormat
         */
        short DataFormat { get; set; }

        /**
         * Get the format string
         */
        String GetDataFormatString();

        /**
         * set the font for this style
         * @param font  a font object Created or retreived from the Workbook object
         * @see Workbook#CreateFont()
         * @see Workbook#GetFontAt(short)
         */

        void SetFont(IFont font);

        /**
         * Gets the index of the font for this style
         * @see Workbook#GetFontAt(short)
         */
        short FontIndex { get; }

        /**
         * get whether the cell's using this style are to be hidden
         * @return hidden - whether the cell using this style should be hidden
         */

        bool IsHidden { get; set; }

        /**
         * get whether the cell's using this style are to be locked
         * @return hidden - whether the cell using this style should be locked
         */

        bool IsLocked { get; set; }


        /**
         * get the type of horizontal alignment for the cell
         * @return align - the type of alignment
         * @see #ALIGN_GENERAL
         * @see #ALIGN_LEFT
         * @see #ALIGN_CENTER
         * @see #ALIGN_RIGHT
         * @see #ALIGN_FILL
         * @see #ALIGN_JUSTIFY
         * @see #ALIGN_CENTER_SELECTION
         */

        HorizontalAlignment Alignment { get; set; }


        /**
         * get whether the text should be wrapped
         * @return wrap text or not
         */

        bool WrapText { get; set; }


        /**
         * get the type of vertical alignment for the cell
         * @return align the type of alignment
         * @see #VERTICAL_TOP
         * @see #VERTICAL_CENTER
         * @see #VERTICAL_BOTTOM
         * @see #VERTICAL_JUSTIFY
         */

        VerticalAlignment VerticalAlignment { get; set; }

        /**
         * get the degree of rotation for the text in the cell
         * @return rotation degrees (between -90 and 90 degrees)
         */

        short Rotation { get; set; }

        /**
         * get the number of spaces to indent the text in the cell
         * @return indent - number of spaces
         */

        short Indention { get; set; }

        /**
         * get the type of border to use for the left border of the cell
         * @return border type
         * @see #BORDER_NONE
         * @see #BORDER_THIN
         * @see #BORDER_MEDIUM
         * @see #BORDER_DASHED
         * @see #BORDER_DOTTED
         * @see #BORDER_THICK
         * @see #BORDER_DOUBLE
         * @see #BORDER_HAIR
         * @see #BORDER_MEDIUM_DASHED
         * @see #BORDER_DASH_DOT
         * @see #BORDER_MEDIUM_DASH_DOT
         * @see #BORDER_DASH_DOT_DOT
         * @see #BORDER_MEDIUM_DASH_DOT_DOT
         * @see #BORDER_SLANTED_DASH_DOT
         */

        BorderStyle BorderLeft { get; set; }


        /**
         * get the type of border to use for the right border of the cell
         * @return border type
         * @see #BORDER_NONE
         * @see #BORDER_THIN
         * @see #BORDER_MEDIUM
         * @see #BORDER_DASHED
         * @see #BORDER_DOTTED
         * @see #BORDER_THICK
         * @see #BORDER_DOUBLE
         * @see #BORDER_HAIR
         * @see #BORDER_MEDIUM_DASHED
         * @see #BORDER_DASH_DOT
         * @see #BORDER_MEDIUM_DASH_DOT
         * @see #BORDER_DASH_DOT_DOT
         * @see #BORDER_MEDIUM_DASH_DOT_DOT
         * @see #BORDER_SLANTED_DASH_DOT
         */

        BorderStyle BorderRight { get; set; }


        /**
         * get the type of border to use for the top border of the cell
         * @return border type
         * @see #BORDER_NONE
         * @see #BORDER_THIN
         * @see #BORDER_MEDIUM
         * @see #BORDER_DASHED
         * @see #BORDER_DOTTED
         * @see #BORDER_THICK
         * @see #BORDER_DOUBLE
         * @see #BORDER_HAIR
         * @see #BORDER_MEDIUM_DASHED
         * @see #BORDER_DASH_DOT
         * @see #BORDER_MEDIUM_DASH_DOT
         * @see #BORDER_DASH_DOT_DOT
         * @see #BORDER_MEDIUM_DASH_DOT_DOT
         * @see #BORDER_SLANTED_DASH_DOT
         */

        BorderStyle BorderTop { get; set; }


        /**
         * get the type of border to use for the bottom border of the cell
         * @return border type
         * @see #BORDER_NONE
         * @see #BORDER_THIN
         * @see #BORDER_MEDIUM
         * @see #BORDER_DASHED
         * @see #BORDER_DOTTED
         * @see #BORDER_THICK
         * @see #BORDER_DOUBLE
         * @see #BORDER_HAIR
         * @see #BORDER_MEDIUM_DASHED
         * @see #BORDER_DASH_DOT
         * @see #BORDER_MEDIUM_DASH_DOT
         * @see #BORDER_DASH_DOT_DOT
         * @see #BORDER_MEDIUM_DASH_DOT_DOT
         * @see #BORDER_SLANTED_DASH_DOT
         */
        BorderStyle BorderBottom { get; set; }


        /**
         * get the color to use for the left border
         */
        short LeftBorderColor { get; set; }


        /**
         * get the color to use for the left border
         * @return the index of the color defInition
         */
        short RightBorderColor { get; set; }


        /**
         * get the color to use for the top border
         * @return hhe index of the color defInition
         */
        short TopBorderColor { get; set; }


        /**
         * get the color to use for the left border
         * @return the index of the color defInition
         */
        short BottomBorderColor { get; set; }


        /**
         * get the fill pattern (??) - set to 1 to fill with foreground color
         * @return fill pattern
         */

        FillPattern FillPattern { get; set; }

        /**
         * get the background fill color
         * @return fill color
         */
        short FillBackgroundColor { get; set; }


        /**
         * get the foreground fill color
         * @return fill color
         */
        short FillForegroundColor { get; set; }

        /**
         * Clones all the style information from another
         *  CellStyle, onto this one. This 
         *  CellStyle will then have all the same
         *  properties as the source, but the two may
         *  be edited independently.
         * Any stylings on this CellStyle will be lost! 
         *  
         * The source CellStyle could be from another
         *  Workbook if you like. This allows you to
         *  copy styles from one Workbook to another.
         *
         * However, both of the CellStyles will need
         *  to be of the same type (HSSFCellStyle or
         *  XSSFCellStyle)
         */
        void CloneStyleFrom(ICellStyle source);

        IFont GetFont(IWorkbook parentWorkbook);

        /// <summary>
        /// Gets or sets the color to use for the diagional border
        /// </summary>
        /// <value>The index of the color definition.</value>
        short BorderDiagonalColor { get; set; }

        /// <summary>
        /// Gets or sets the line type  to use for the diagional border
        /// </summary>
        /// <value>The line type.</value>
        BorderStyle BorderDiagonalLineStyle { get; set; }

        /// <summary>
        /// Gets or sets the type of diagional border
        /// </summary>.
        /// <value>The border diagional type.</value>
        BorderDiagonal BorderDiagonal { get; set; }

        /**
         * Gets the color object representing the current
         *  background fill, resolving indexes using
         *  the supplied workbook.
         * This will work for both indexed and rgb
         *  defined colors. 
         */
        IColor FillBackgroundColorColor { get; }
        /**
         * Gets the color object representing the current
         *  foreground fill, resolving indexes using
         *  the supplied workbook.
         * This will work for both indexed and rgb
         *  defined colors. 
         */
        IColor FillForegroundColorColor { get; }
    }
}
