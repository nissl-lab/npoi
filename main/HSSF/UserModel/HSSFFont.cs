/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


using NPOI.HSSF.Util;

namespace NPOI.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.Record;
    using NPOI.SS.UserModel;

    /// <summary>
    /// Represents a Font used in a workbook.
    /// @version 1.0-pre
    /// @author  Andrew C. Oliver
    /// </summary>
    public class HSSFFont:NPOI.SS.UserModel.IFont
    {
        public const String FONT_ARIAL = "Arial";

        private FontRecord font;
        private short index;

        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFFont"/> class.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <param name="rec">The record.</param>
        public HSSFFont(short index, FontRecord rec)
        {
            font = rec;
            this.index = index;
        }


        /// <summary>
        /// Get the name for the font (i.e. Arial)
        /// </summary>
        /// <value>the name of the font to use</value>
        public String FontName
        {
            get { return font.FontName; }
            set
            {
                font.FontName = value;
            }
        }

        /// <summary>
        /// Get the index within the HSSFWorkbook (sequence within the collection of Font objects)
        /// </summary>
        /// <value>Unique index number of the Underlying record this Font represents (probably you don't care
        /// Unless you're comparing which one is which)</value>
        public short Index
        {
            get { return index; }
        }



        /// <summary>
        /// Get or sets the font height in Unit's of 1/20th of a point.  Maybe you might want to
        /// use the GetFontHeightInPoints which matches to the familiar 10, 12, 14 etc..
        /// </summary>
        /// <value>height in 1/20ths of a point.</value>
        public double FontHeight
        {
            get { return font.FontHeight; }
            set { font.FontHeight = (short)value; }
        }

        /// <summary>
        /// Gets or sets the font height in points.
        /// </summary>
        /// <value>height in the familiar Unit of measure - points.</value>
        public short FontHeightInPoints
        {
            get { return (short)(font.FontHeight / 20); }
            set { font.FontHeight=(short)(value * 20); }
        }

        /// <summary>
        /// Gets or sets whether to use italics or not
        /// </summary>
        /// <value><c>true</c> if this instance is italic; otherwise, <c>false</c>.</value>
        public bool IsItalic
        {
            get { return font.IsItalic; }
            set { font.IsItalic=value; }
        }

        /// <summary>
        /// Get whether to use a strikeout horizontal line through the text or not
        /// </summary>
        /// <value>
        /// strikeout or not
        /// </value>
        public bool IsStrikeout
        {
            get { return font.IsStrikeout; }
            set { font.IsStrikeout=value; }
        }

        /// <summary>
        /// Gets or sets the color for the font.
        /// </summary>
        /// <value>The color to use.</value>
        public short Color
        {
            get { return font.ColorPaletteIndex; }
            set { font.ColorPaletteIndex=value; }
        }

        /// <summary>
        /// get the color value for the font
        /// </summary>
        /// <param name="wb">HSSFWorkbook</param>
        /// <returns></returns>
        public HSSFColor GetHSSFColor(HSSFWorkbook wb)
        {
            HSSFPalette pallette = wb.GetCustomPalette();
            return pallette.GetColor(Color);
        }

        /// <summary>
        /// Gets or sets the boldness to use
        /// </summary>
        /// <value>The boldweight.</value>
        public short Boldweight
        {
            get { return font.BoldWeight; }
            set { font.BoldWeight = value; }
        }
        /**
         * get or set if the font bold style
         */
        public bool IsBold
        {
            get
            {
                return font.BoldWeight == (short)FontBoldWeight.Bold;
            }
            set
            {
                if (value)
                    font.BoldWeight = (short)FontBoldWeight.Bold;
                else
                    font.BoldWeight = (short)FontBoldWeight.Normal;
            }
        }

        /// <summary>
        /// Gets or sets normal,base or subscript.
        /// </summary>
        /// <value>offset type to use (none,base,sub)</value>
        public FontSuperScript TypeOffset
        {
            get { return font.SuperSubScript; }
            set { font.SuperSubScript = value; }
        }


        /// <summary>
        /// Gets or sets the type of text Underlining to use
        /// </summary>
        /// <value>The Underlining type.</value>
        public FontUnderlineType Underline
        {
            get { return font.Underline; }
            set { font.Underline = value; }
        }


        /// <summary>
        /// Gets or sets the char set to use.
        /// </summary>
        /// <value>The char set.</value>
        public short Charset
        {
            get { return font.Charset; }
            set { font.Charset = (byte)value; }
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            return "NPOI.HSSF.UserModel.HSSFFont{" +
                     font +
                    "}";
        }

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="T:System.Object"/>.
        /// </returns>
        public override int GetHashCode()
        {
            int prime = 31;
            int result = 1;
            result = prime * result + ((font == null) ? 0 : font.GetHashCode());
            result = prime * result + index;
            return result;
        }

        /// <summary>
        /// Determines whether the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <param name="obj">The <see cref="T:System.Object"/> to compare with the current <see cref="T:System.Object"/>.</param>
        /// <returns>
        /// true if the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>; otherwise, false.
        /// </returns>
        /// <exception cref="T:System.NullReferenceException">
        /// The <paramref name="obj"/> parameter is null.
        /// </exception>
        public override bool Equals(Object obj)
        {
            if (this == obj) return true;
            if (obj == null) return false;
            if (obj is HSSFFont)
            {
                HSSFFont other = (HSSFFont)obj;
                if (font == null)
                {
                    if (other.font != null)
                        return false;
                }
                else if (!font.Equals(other.font))
                    return false;
                if (index != other.index)
                    return false;
                return true;
            }
            return false;
        }
    }
}