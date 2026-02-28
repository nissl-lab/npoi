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
    using Org.BouncyCastle.Utilities;
    using System;
    using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;

    public interface ICellStyle
    {
        /// <summary>
        /// the Cell should be auto-sized to shrink to fit if the text is too long
        /// </summary>
        bool ShrinkToFit { get; set; }

        /// <summary>
        /// get the index within the Workbook(sequence within the collection of ExtnededFormat objects)
        /// </summary>
        short Index { get; }

        /// <summary>
        /// get or set the index of the format
        /// </summary>
        short DataFormat { get; set; }
        /// <summary>
        /// Get the format string
        /// </summary>
        /// <returns></returns>
        String GetDataFormatString();

        /// <summary>
        /// set the font for this style
        /// </summary>
        /// <param name="font">a font object created or retreived from the Workbook object</param>
        void SetFont(IFont font);

        /// <summary>
        /// Gets the index of the font for this style
        /// </summary>
        short FontIndex { get; }

        /// <summary>
        /// get or set whether the cell's using this style are to be hidden
        /// </summary>
        bool IsHidden { get; set; }

        /// <summary>
        /// get or set whether the cell's using this style are to be locked
        /// </summary>
        bool IsLocked { get; set; }

        /// <summary>
        /// Turn on or off "Quote Prefix" or "123 Prefix" for the style,
        /// which is used to tell Excel that the thing which looks like
        /// a number or a formula shouldn't be treated as on.
        /// Turning this on is somewhat (but not completely, 
        /// like prefixing the cell value with a ' in Excel
        /// </summary>
        bool IsQuotePrefixed { get; set; }

        /// <summary>
        /// get or set the type of horizontal alignment for the cell
        /// </summary>
        HorizontalAlignment Alignment { get; set; }

        /// <summary>
        /// get or set whether the text should be wrapped
        /// </summary>
        bool WrapText { get; set; }

        /// <summary>
        /// get or set the type of vertical alignment for the cell
        /// </summary>
        VerticalAlignment VerticalAlignment { get; set; }

        /// <summary>
        /// get or set the degree of rotation for the text in the cell
        /// </summary>
        /// <remarks>
        /// Note: HSSF uses values from -90 to 90 degrees, whereas XSSF 
        /// uses values from 0 to 180 degrees. The implementations of this method will map between these two value-ranges
        /// accordingly, however the corresponding getter is returning values in the range mandated by the current type
        /// of Excel file-format that this CellStyle is applied to.
        /// </remarks>
        short Rotation { get; set; }

        /// <summary>
        /// get the number of spaces to indent the text in the cell
        /// </summary>
        short Indention { get; set; }

        /// <summary>
        /// get or set the type of border to use for the left border of the cell
        /// </summary>
        BorderStyle BorderLeft { get; set; }

        /// <summary>
        /// get or set the type of border to use for the right border of the cell
        /// </summary>
        BorderStyle BorderRight { get; set; }

        /// <summary>
        /// get or set the type of border to use for the top border of the cell
        /// </summary>
        BorderStyle BorderTop { get; set; }

        /// <summary>
        /// get or set the type of border to use for the bottom border of the cell
        /// </summary>
        BorderStyle BorderBottom { get; set; }

        /// <summary>
        /// get or set the color to use for the left border
        /// </summary>
        short LeftBorderColor { get; set; }

        /// <summary>
        /// get or set the color to use for the left border
        /// </summary>
        short RightBorderColor { get; set; }

        /// <summary>
        /// get or set the color to use for the top border
        /// </summary>
        short TopBorderColor { get; set; }

        /// <summary>
        /// get or set the color to use for the left border
        /// </summary>
        short BottomBorderColor { get; set; }

        /// <summary>
        /// get or set the fill pattern(??) - set to 1 to fill with foreground color
        /// </summary>
        FillPattern FillPattern { get; set; }

        /// <summary>
        /// get or set the background fill color
        /// </summary>
        short FillBackgroundColor { get; set; }

        /// <summary>
        /// get or set the foreground fill color
        /// </summary>
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

        /// <summary>
        /// Gets the color object representing the current
        /// background fill, resolving indexes using the supplied workbook.
        /// This will work for both indexed and rgb defined colors.
        /// </summary>
        IColor FillBackgroundColorColor { get; }
        /// <summary>
        ///  Gets the color object representing the current 
        /// foreground fill, resolving indexes using the supplied workbook.
        /// This will work for both indexed and rgb defined colors. 
        /// </summary>
        IColor FillForegroundColorColor { get; }
        /// <summary>
        /// Get or set the reading order, for RTL/LTR ordering of the text.
        /// </summary>
        ReadingOrder ReadingOrder { get; set; }
    }
}
