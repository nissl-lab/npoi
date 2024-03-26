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
        /// <summary>
        /// get the index within the Workbook (sequence within the collection of ExtnededFormat objects)
        /// </summary>
        /// <return>unique index number of the underlying record this style represents (probably you don't care
        /// unless you're comparing which one is which)
        /// </return>

        short Index { get; }

        /// <summary>
        /// get the index of the format
        /// </summary>
        /// <see cref="IDataFormat" />
        short DataFormat { get; set; }

        /// <summary>
        /// Get the format string
        /// </summary>
        String GetDataFormatString();

        /// <summary>
        /// set the font for this style
        /// </summary>
        /// <param name="font"> a font object Created or retreived from the Workbook object</param>
        /// <see cref="IWorkbook.CreateFont()" />
        /// <see cref="IWorkbook.GetFontAt(short)" />

        void SetFont(IFont font);

        /// <summary>
        /// Gets the index of the font for this style
        /// </summary>
        /// <see cref="IWorkbook.GetFontAt(short)" />
        short FontIndex { get; }

        /// <summary>
        /// get whether the cell's using this style are to be hidden
        /// </summary>
        /// <return>hidden - whether the cell using this style should be hidden</return>

        bool IsHidden { get; set; }

        /// <summary>
        /// get whether the cell's using this style are to be locked
        /// </summary>
        /// <return>hidden - whether the cell using this style should be locked</return>

        bool IsLocked { get; set; }

        /// <summary>
        /// Turn on or off "Quote Prefix" or "123 Prefix" for the style,
        /// which is used to tell Excel that the thing which looks like
        /// a number or a formula shouldn't be treated as on.
        /// Turning this on is somewhat (but not completely, see {@link IgnoredErrorType})
        /// like prefixing the cell value with a ' in Excel
        /// </summary>
        bool IsQuotePrefixed { get; set; }

        /// <summary>
        /// get the type of horizontal alignment for the cell
        /// </summary>
        /// <return>align - the type of alignment</return>
        /// <see cref="ALIGN_GENERAL" />
        /// <see cref="ALIGN_LEFT" />
        /// <see cref="ALIGN_CENTER" />
        /// <see cref="ALIGN_RIGHT" />
        /// <see cref="ALIGN_FILL" />
        /// <see cref="ALIGN_JUSTIFY" />
        /// <see cref="ALIGN_CENTER_SELECTION" />

        HorizontalAlignment Alignment { get; set; }


        /// <summary>
        /// get whether the text should be wrapped
        /// </summary>
        /// <return>wrap text or not</return>

        bool WrapText { get; set; }


        /// <summary>
        /// get the type of vertical alignment for the cell
        /// </summary>
        /// <return>align the type of alignment</return>
        /// <see cref="VERTICAL_TOP" />
        /// <see cref="VERTICAL_CENTER" />
        /// <see cref="VERTICAL_BOTTOM" />
        /// <see cref="VERTICAL_JUSTIFY" />

        VerticalAlignment VerticalAlignment { get; set; }

        /// <summary>
        /// <para>
        /// get the degree of rotation for the text in the cell
        /// </para>
        /// <para>
        /// Note: HSSF uses values from -90 to 90 degrees, whereas XSSF
        /// uses values from 0 to 180 degrees. The implementations of this method will map between these two value-ranges
        /// accordingly, however the corresponding getter is returning values in the range mandated by the current type
        /// of Excel file-format that this CellStyle is applied to.
        /// </para>
        /// </summary>
        /// <return>rotation degrees (between -90 and 90 degrees)</return>

        short Rotation { get; set; }

        /// <summary>
        /// get the number of spaces to indent the text in the cell
        /// </summary>
        /// <return>indent - number of spaces</return>

        short Indention { get; set; }

        /// <summary>
        /// get the type of border to use for the left border of the cell
        /// </summary>
        /// <return>border type</return>
        /// <see cref="BorderStyle.None" />
        /// <see cref="BORDER_THIN" />
        /// <see cref="BORDER_MEDIUM" />
        /// <see cref="BORDER_DASHED" />
        /// <see cref="BORDER_DOTTED" />
        /// <see cref="BORDER_THICK" />
        /// <see cref="BORDER_DOUBLE" />
        /// <see cref="BORDER_HAIR" />
        /// <see cref="BORDER_MEDIUM_DASHED" />
        /// <see cref="BORDER_DASH_DOT" />
        /// <see cref="BORDER_MEDIUM_DASH_DOT" />
        /// <see cref="BORDER_DASH_DOT_DOT" />
        /// <see cref="BORDER_MEDIUM_DASH_DOT_DOT" />
        /// <see cref="BORDER_SLANTED_DASH_DOT" />

        BorderStyle BorderLeft { get; set; }


        /// <summary>
        /// get the type of border to use for the right border of the cell
        /// </summary>
        /// <return>border type</return>
        /// <see cref="BORDER_NONE" />
        /// <see cref="BORDER_THIN" />
        /// <see cref="BORDER_MEDIUM" />
        /// <see cref="BORDER_DASHED" />
        /// <see cref="BORDER_DOTTED" />
        /// <see cref="BORDER_THICK" />
        /// <see cref="BORDER_DOUBLE" />
        /// <see cref="BORDER_HAIR" />
        /// <see cref="BORDER_MEDIUM_DASHED" />
        /// <see cref="BORDER_DASH_DOT" />
        /// <see cref="BORDER_MEDIUM_DASH_DOT" />
        /// <see cref="BORDER_DASH_DOT_DOT" />
        /// <see cref="BORDER_MEDIUM_DASH_DOT_DOT" />
        /// <see cref="BORDER_SLANTED_DASH_DOT" />

        BorderStyle BorderRight { get; set; }


        /// <summary>
        /// get the type of border to use for the top border of the cell
        /// </summary>
        /// <return>border type</return>
        /// <see cref="BORDER_NONE" />
        /// <see cref="BORDER_THIN" />
        /// <see cref="BORDER_MEDIUM" />
        /// <see cref="BORDER_DASHED" />
        /// <see cref="BORDER_DOTTED" />
        /// <see cref="BORDER_THICK" />
        /// <see cref="BORDER_DOUBLE" />
        /// <see cref="BORDER_HAIR" />
        /// <see cref="BORDER_MEDIUM_DASHED" />
        /// <see cref="BORDER_DASH_DOT" />
        /// <see cref="BORDER_MEDIUM_DASH_DOT" />
        /// <see cref="BORDER_DASH_DOT_DOT" />
        /// <see cref="BORDER_MEDIUM_DASH_DOT_DOT" />
        /// <see cref="BORDER_SLANTED_DASH_DOT" />

        BorderStyle BorderTop { get; set; }


        /// <summary>
        /// get the type of border to use for the bottom border of the cell
        /// </summary>
        /// <return>border type</return>
        /// <see cref="BORDER_NONE" />
        /// <see cref="BORDER_THIN" />
        /// <see cref="BORDER_MEDIUM" />
        /// <see cref="BORDER_DASHED" />
        /// <see cref="BORDER_DOTTED" />
        /// <see cref="BORDER_THICK" />
        /// <see cref="BORDER_DOUBLE" />
        /// <see cref="BORDER_HAIR" />
        /// <see cref="BORDER_MEDIUM_DASHED" />
        /// <see cref="BORDER_DASH_DOT" />
        /// <see cref="BORDER_MEDIUM_DASH_DOT" />
        /// <see cref="BORDER_DASH_DOT_DOT" />
        /// <see cref="BORDER_MEDIUM_DASH_DOT_DOT" />
        /// <see cref="BORDER_SLANTED_DASH_DOT" />
        BorderStyle BorderBottom { get; set; }


        /// <summary>
        /// get the color to use for the left border
        /// </summary>
        short LeftBorderColor { get; set; }


        /// <summary>
        /// get the color to use for the left border
        /// </summary>
        /// <return>the index of the color defInition</return>
        short RightBorderColor { get; set; }


        /// <summary>
        /// get the color to use for the top border
        /// </summary>
        /// <return>hhe index of the color defInition</return>
        short TopBorderColor { get; set; }


        /// <summary>
        /// get the color to use for the left border
        /// </summary>
        /// <return>the index of the color defInition</return>
        short BottomBorderColor { get; set; }


        /// <summary>
        /// get the fill pattern (??) - set to 1 to fill with foreground color
        /// </summary>
        /// <return>fill pattern</return>

        FillPattern FillPattern { get; set; }

        /// <summary>
        /// get the background fill color
        /// </summary>
        /// <return>fill color</return>
        short FillBackgroundColor { get; set; }


        /// <summary>
        /// get the foreground fill color
        /// </summary>
        /// <return>fill color</return>
        short FillForegroundColor { get; set; }

        /// <summary>
        /// <para>
        /// Clones all the style information from another
        ///  CellStyle, onto this one. This
        ///  CellStyle will then have all the same
        ///  properties as the source, but the two may
        ///  be edited independently.
        /// Any stylings on this CellStyle will be lost!
        /// </para>
        /// <para>
        /// The source CellStyle could be from another
        ///  Workbook if you like. This allows you to
        ///  copy styles from one Workbook to another.
        /// </para>
        /// <para>
        /// However, both of the CellStyles will need
        ///  to be of the same type (HSSFCellStyle or
        ///  XSSFCellStyle)
        /// </para>
        /// </summary>
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
        ///  background fill, resolving indexes using
        ///  the supplied workbook.
        /// This will work for both indexed and rgb
        ///  defined colors.
        /// </summary>
        IColor FillBackgroundColorColor { get; }
        /// <summary>
        /// Gets the color object representing the current
        ///  foreground fill, resolving indexes using
        ///  the supplied workbook.
        /// This will work for both indexed and rgb
        ///  defined colors.
        /// </summary>
        IColor FillForegroundColorColor { get; }
    }
}
