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
        /**
         * not underlined
         */
        None = 0,
        /**
         * single (normal) underline
         */
        Single = 1,
        /**
         * double underlined
         */
        Double = 2,
        /**
         * accounting style single underline
         */
        SingleAccounting = 0x21,
        /**
         * accounting style double underline
         */
        DoubleAccounting = 0x22
    }

    public enum FontSuperScript:short
    { 
    
        /**
         * no type Offsetting (not super or subscript)
         */

        None = 0,

        /**
         * superscript
         */

        Super = 1,

        /**
         * subscript
         */

        Sub = 2,
    }

    public enum FontColor:short
    {
        /// <summary>
        /// Allow accessing the Initial value.
        /// </summary>
        None = 0,

        /**
         * normal type of black color.
         */

        Normal = 0x7fff,

        /**
         * Dark Red color
         */

        Red = 0xa,
    }

    public enum FontBoldWeight:short
    {
        /// <summary>
        /// Allow accessing the Initial value.
        /// </summary>
        None = 0,

        /**
         * Normal boldness (not bold)
         */

        Normal = 0x190,

        /**
         * Bold boldness (bold)
         */

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

        /**
         * get the color for the font
         * @return color to use
         * @see #COLOR_NORMAL
         * @see #COLOR_RED
         * @see NPOI.HSSF.usermodel.HSSFPalette#GetColor(short)
         */
        short Color { get; set; }

        /// <summary>
        ///  get type of text underlining to use
        /// </summary>
        FontSuperScript TypeOffset { get; set; }

        /// <summary>
        /// get type of text underlining to use
        /// </summary>
        FontUnderlineType Underline { get; set; }

        /**
         * get character-set to use.
         * @return character-set
         * @see #ANSI_CHARSET
         * @see #DEFAULT_CHARSET
         * @see #SYMBOL_CHARSET
         */
        short Charset { get; set; }

        /**
         * get the index within the XSSFWorkbook (sequence within the collection of Font objects)
         * 
         * @return unique index number of the underlying record this Font represents (probably you don't care
         *  unless you're comparing which one is which)
         */
        short Index { get; }

        short Boldweight { get; set; }

        bool IsBold { get; set; }

        /// <summary>
        /// Copies the style from another font into this one
        /// </summary>
        void CloneStyleFrom(IFont src);
    }
}