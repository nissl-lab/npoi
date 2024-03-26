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

    public enum FontUnderlineType : byte
    {
        /// <summary>
        /// not underlined
        /// </summary>
        None = 0,
        /// <summary>
        /// single (normal) underline
        /// </summary>
        Single = 1,
        /// <summary>
        /// double underlined
        /// </summary>
        Double = 2,
        /// <summary>
        /// accounting style single underline
        /// </summary>
        SingleAccounting = 0x21,
        /// <summary>
        /// accounting style double underline
        /// </summary>
        DoubleAccounting = 0x22
    }

    public enum FontSuperScript : short
    {

        /// <summary>
        /// no type Offsetting (not super or subscript)
        /// </summary>

        None = 0,

        /// <summary>
        /// superscript
        /// </summary>

        Super = 1,

        /// <summary>
        /// subscript
        /// </summary>

        Sub = 2,
    }

    public enum FontColor : short
    {
        /// <summary>
        /// Allow accessing the Initial value.
        /// </summary>
        None = 0,

        /// <summary>
        /// normal type of black color.
        /// </summary>

        Normal = 0x7fff,

        /// <summary>
        /// Dark Red color
        /// </summary>

        Red = 0xa,
    }
    [Obsolete("deprecated POI 3.15 beta 2. Boldweight constants no longer needed due to IsBold property")]
    public enum FontBoldWeight : short
    {
        /// <summary>
        /// Allow accessing the Initial value.
        /// </summary>
        None = 0,

        /// <summary>
        /// Normal boldness (not bold)
        /// </summary>

        Normal = 0x190,

        /// <summary>
        /// Bold boldness (bold)
        /// </summary>

        Bold = 0x2bc,
    }


    public interface IFont
    {

        /// <summary>
        /// get the name for the font (i.e. Arial)
        /// </summary>
        String FontName { get; set; }

        /// <summary>
        ///  Get the font height in unit's of 1/20th of a point.
        /// </summary>
        /// <remarks>
        /// Maybe you might want to use the GetFontHeightInPoints which matches to the familiar 10, 12, 14 etc..
        /// </remarks>
        /// <see cref="FontHeightInPoints"/>
        double FontHeight { get; set; }

        /// <summary>
        /// Get the font height in points.
        /// </summary>
        /// <remarks>
        /// This will return the same font height that is shown in Excel, such as 10 or 14 or 28.
        /// </remarks>
        /// <see cref="FontHeight"/>
        double FontHeightInPoints { get; set; }

        /// <summary>
        /// get whether to use italics or not
        /// </summary>
        bool IsItalic { get; set; }

        /// <summary>
        /// get whether to use a strikeout horizontal line through the text or not
        /// </summary>
        bool IsStrikeout { get; set; }

        /// <summary>
        /// get the color for the font
        /// </summary>
        /// <returns>color to use</returns>
        /// <see cref="COLOR_NORMAL" />
        /// <see cref="COLOR_RED" />
        /// <see cref="NPOI.HSSF.UserModel.HSSFPalette.GetColor(short)" />
        short Color { get; set; }

        /// <summary>
        ///  get type of text underlining to use
        /// </summary>
        FontSuperScript TypeOffset { get; set; }

        /// <summary>
        /// get type of text underlining to use
        /// </summary>
        FontUnderlineType Underline { get; set; }

        /// <summary>
        /// get character-set to use.
        /// </summary>
        /// <value>ANSI_CHARSET,DEFAULT_CHARSET,SYMBOL_CHARSET </value>
        short Charset { get; set; }

        /// <summary>
        /// get the index within the Workbook (sequence within the collection of Font objects)
        /// </summary>
        short Index { get; }
        [Obsolete("deprecated POI 3.15 beta 2. Use IsBold instead.")]
        short Boldweight { get; set; }

        bool IsBold { get; set; }

        /// <summary>
        /// Copies the style from another font into this one
        /// </summary>
        void CloneStyleFrom(IFont src);
    }
}