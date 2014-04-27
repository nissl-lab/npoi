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

namespace NPOI.HSSF.UserModel
{
    using System;
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Util;
    using NPOI.Util;


    /// <summary>
    /// Represents a workbook color palette.
    /// Internally, the XLS format refers to colors using an offset into the palette
    /// record.  Thus, the first color in the palette has the index 0x8, the second
    /// has the index 0x9, etc. through 0x40
    /// @author Brian Sanders (bsanders at risklabs dot com)
    /// </summary>
    public class HSSFPalette
    {
        private PaletteRecord palette;

        public HSSFPalette(PaletteRecord palette)
        {
            this.palette = palette;
        }

        /// <summary>
        /// Retrieves the color at a given index
        /// </summary>
        /// <param name="index">the palette index, between 0x8 to 0x40 inclusive.</param>
        /// <returns>the color, or null if the index Is not populated</returns>
        public HSSFColor GetColor(short index)
        {
            //Handle the special AUTOMATIC case
            if (index == HSSFColor.Automatic.Index)
                return HSSFColor.Automatic.GetInstance();
            else
            {
                byte[] b = palette.GetColor(index);
                if (b != null)
                {
                    return new CustomColor(index, b);
                }
            }
            return null;
        }

        /// <summary>
        /// Finds the first occurance of a given color
        /// </summary>
        /// <param name="red">the RGB red component, between 0 and 255 inclusive</param>
        /// <param name="green">the RGB green component, between 0 and 255 inclusive</param>
        /// <param name="blue">the RGB blue component, between 0 and 255 inclusive</param>
        /// <returns>the color, or null if the color does not exist in this palette</returns>
        public HSSFColor FindColor(byte red, byte green, byte blue)
        {
            byte[] b = palette.GetColor(PaletteRecord.FIRST_COLOR_INDEX);
            for (short i = (short)PaletteRecord.FIRST_COLOR_INDEX; b != null;
                b = palette.GetColor(++i))
            {
                if (b[0] == red && b[1] == green && b[2] == blue)
                {
                    return new CustomColor(i, b);
                }
            }
            return null;
        }

        /// <summary>
        /// Finds the closest matching color in the custom palette.  The
        /// method for Finding the distance between the colors Is fairly
        /// primative.
        /// </summary>
        /// <param name="red">The red component of the color to match.</param>
        /// <param name="green">The green component of the color to match.</param>
        /// <param name="blue">The blue component of the color to match.</param>
        /// <returns>The closest color or null if there are no custom
        /// colors currently defined.</returns>
        public HSSFColor FindSimilarColor(byte red, byte green, byte blue)
        {
            HSSFColor result = null;
            int minColorDistance = int.MaxValue;
            byte[] b = palette.GetColor(PaletteRecord.FIRST_COLOR_INDEX);
            for (short i = (short)PaletteRecord.FIRST_COLOR_INDEX; b != null;
                b = palette.GetColor(++i))
            {
                int colorDistance = Math.Abs(red - b[0]) +
                    Math.Abs(green - b[1]) + Math.Abs(blue - b[2]);
                if (colorDistance < minColorDistance)
                {
                    minColorDistance = colorDistance;
                    result = GetColor(i);
                }
            }
            return result;
        }

        /// <summary>
        /// Sets the color at the given offset
        /// </summary>
        /// <param name="index">the palette index, between 0x8 to 0x40 inclusive</param>
        /// <param name="red">the RGB red component, between 0 and 255 inclusive</param>
        /// <param name="green">the RGB green component, between 0 and 255 inclusive</param>
        /// <param name="blue">the RGB blue component, between 0 and 255 inclusive</param>
        public void SetColorAtIndex(short index, byte red, byte green, byte blue)
        {
            palette.SetColor(index, red, green, blue);
        }

        /// <summary>
        /// Adds a new color into an empty color slot.
        /// </summary>
        /// <param name="red">The red component</param>
        /// <param name="green">The green component</param>
        /// <param name="blue">The blue component</param>
        /// <returns>The new custom color.</returns>
        public HSSFColor AddColor(byte red, byte green, byte blue)
        {
            byte[] b = palette.GetColor(PaletteRecord.FIRST_COLOR_INDEX);
            short i;
            for (i = (short)PaletteRecord.FIRST_COLOR_INDEX; i < PaletteRecord.STANDARD_PALETTE_SIZE + PaletteRecord.FIRST_COLOR_INDEX; b = palette.GetColor(++i))
            {
                if (b == null)
                {
                    SetColorAtIndex(i, red, green, blue);
                    return GetColor(i);
                }
            }
            throw new Exception("Could not Find free color index");
        }

        /// <summary>
        /// user custom color
        /// </summary>
        private class CustomColor : HSSFColor
        {
            private short byteOffset;
            private byte red;
            private byte green;
            private byte blue;

            /// <summary>
            /// Initializes a new instance of the <see cref="CustomColor"/> class.
            /// </summary>
            /// <param name="byteOffset">The byte offset.</param>
            /// <param name="colors">The colors.</param>
            public CustomColor(short byteOffset, byte[] colors): this(byteOffset, colors[0], colors[1], colors[2])
            {
               
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="CustomColor"/> class.
            /// </summary>
            /// <param name="byteOffset">The byte offset.</param>
            /// <param name="red">The red.</param>
            /// <param name="green">The green.</param>
            /// <param name="blue">The blue.</param>
            public CustomColor(short byteOffset, byte red, byte green, byte blue)
            {
                this.byteOffset = byteOffset;
                this.red = red;
                this.green = green;
                this.blue = blue;
            }

            /// <summary>
            /// Gets index to the standard palette
            /// </summary>
            /// <value></value>
            public override short Indexed
            {
                get
                {
                    return byteOffset;
                }
            }

            /// <summary>
            /// Gets triplet representation like that in Excel
            /// </summary>
            /// <value></value>
            public override byte[] GetTriplet()
            {
                    return new byte[]
                    {
                        (byte)(red   & 0xff),
                        (byte)(green & 0xff),
                        (byte)(blue  & 0xff)
                    };
            }

            /// <summary>
            /// Gets a hex string exactly like a gnumeric triplet
            /// </summary>
            /// <value></value>
            public override String GetHexString()
            {
                    StringBuilder sb = new StringBuilder();
                    sb.Append(GetGnumericPart(red));
                    sb.Append(':');
                    sb.Append(GetGnumericPart(green));
                    sb.Append(':');
                    sb.Append(GetGnumericPart(blue));
                    return sb.ToString();
            }

            /// <summary>
            /// Gets the gnumeric part.
            /// </summary>
            /// <param name="color">The color.</param>
            /// <returns></returns>
            private String GetGnumericPart(byte color)
            {
                String s;
                if (color == 0)
                {
                    s = "0";
                }
                else
                {
                    int c = color & 0xff; //as Unsigned
                    c = (c << 8) | c; //pad to 16-bit
                    s = StringUtil.ToHexString(c).ToUpper();
                    while (s.Length < 4)
                    {
                        s = "0" + s;
                    }
                }
                return s;
            }
        }
    }
}