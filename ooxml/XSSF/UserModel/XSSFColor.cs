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
using NPOI.OpenXmlFormats;
using System;
using NPOI.HSSF.Util;
using System.Text;
using NPOI.Util;
using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace NPOI.XSSF.UserModel
{

    /**
     * Represents a color in SpreadsheetML
     */
    public class XSSFColor : ExtendedColor
    {

        private CT_Color ctColor;

        /**
         * Create an instance of XSSFColor from the supplied XML bean
         */
        public XSSFColor(CT_Color color)
        {
            this.ctColor = color;
        }

        /**
         * Create an new instance of XSSFColor
         */
        public XSSFColor()
        {
            this.ctColor = new CT_Color();
        }

        public XSSFColor(Color clr)
            : this()
        {
            var c = clr.ToPixel<Rgb24>();
            ctColor.SetRgb(c.R, c.G, c.B);
        }

        public XSSFColor(Rgb24 clr)
            : this() {

            ctColor.SetRgb(clr.R, clr.G, clr.B);
        }

        public XSSFColor(byte[] rgb)
            : this()
        {

            ctColor.SetRgb(rgb);
        }

        public XSSFColor(IndexedColors indexedColor)
            : this()
        {
            ctColor.indexed = (uint)indexedColor.Index;
        }
        /// <summary>
        ///A bool value indicating the ctColor is automatic and system ctColor dependent.
        /// </summary>
        public override bool IsAuto
        {
            get
            {
                return ctColor.auto;
            }
            set 
            {
                ctColor.auto = value;
                ctColor.autoSpecified = true;
            }
        }
        public override bool IsIndexed
        {
            get
            {
                return ctColor.IsSetIndexed();
            }
            
        }
        public override bool IsRGB
        {
            get
            {
                return ctColor.IsSetRgb();
            }
        }

        public override bool IsThemed
        {
            get
            {
                return ctColor.IsSetTheme();
            }
        }

        /**
         * A bool value indicating if the ctColor has a alpha or not
         */
        public bool HasAlpha
        {
            get
            {
                if (!ctColor.IsSetRgb()) return false;
                return ctColor.rgb.Length == 4;
            }
            
        }
        /**
         * A bool value indicating if the ctColor has a tint or not
         */
        public bool HasTint
        {
            get
            {
                if (!ctColor.IsSetTint())
                {
                    return false;
                }
                return ctColor.tint != 0;
            }
        }
        public override short Index
        {
            get
            {
                return ctColor.indexedSpecified ? (short)ctColor.indexed : (short)0;
            }
        }
        /**
         * Indexed ctColor value. Only used for backwards compatibility. References a ctColor in indexedColors.
         */
        public short Indexed
        {
            get
            {
                return Index;
            }
            set 
            {
                ctColor.indexed = (uint)value;
                ctColor.indexedSpecified = true;
            }
        }

        [Obsolete("use property RGB")]
        public byte[] GetRgb()
        { 
            return this.RGB;
        }

        protected override byte[] StoredRBG
        {
            get
            {
                return ctColor.rgb;
            }
            
        }
        /**
         * Standard Red Green Blue ctColor value (RGB).
         * If there was an A (Alpha) value, it will be stripped.
         */
        public override byte[] RGB
        {
            get
            {
                byte[] rgb = GetRGBOrARGB();
                if (rgb == null) return null;

                if (rgb.Length == 4)
                {
                    // Need to trim off the alpha
                    byte[] tmp = new byte[3];
                    Array.Copy(rgb, 1, tmp, 0, 3);
                    return tmp;
                }
                else
                {
                    return rgb;
                }
            }
            set
            {
                ctColor.SetRgb((value));
            }
        }
        /**
        * Standard Alpha Red Green Blue ctColor value (ARGB).
        */
        public override byte[] ARGB
        {
            get
            {
                byte[] rgb = GetRGBOrARGB();
                if (rgb == null) return null;

                if (rgb.Length == 3)
                {
                    // Pad with the default Alpha
                    byte[] tmp = new byte[4];
                    unchecked
                    {
                        tmp[0] = (byte)-1;
                    }
                    Array.Copy(rgb, 0, tmp, 1, 3);
                    return tmp;
                }
                else
                {
                    return rgb;
                }
            }
            
        }
        /**
         * Standard Alpha Red Green Blue ctColor value (ARGB).
         */
        [Obsolete("use property ARGB")]
        public byte[] GetARgb()
        {
            return ARGB;
        }


        /**
         * Standard Red Green Blue ctColor value (RGB) with applied tint.
         * Alpha values are ignored.
         */
        public byte[] GetRgbWithTint()
        {
            byte[] rgb = ctColor.GetRgb();
            if (rgb != null)
            {
                if (rgb.Length == 4)
                {
                    byte[] tmp = new byte[3];
                    Array.Copy(rgb, 1, tmp, 0, 3);
                    rgb = tmp;
                }
                for (int i = 0; i < rgb.Length; i++)
                {
                    rgb[i] = ApplyTint(rgb[i] & 0xFF, ctColor.tint);
                }
            }
            return rgb;
        }


        private static byte ApplyTint(int lum, double tint)
        {
            if (tint > 0)
            {
                return (byte)(lum * (1.0 - tint) + (255 - 255 * (1.0 - tint)));
            }
            else if (tint < 0)
            {
                return (byte)(lum * (1 + tint));
            }
            else
            {
                return (byte)lum;
            }
        }

        /**
         * Standard Alpha Red Green Blue ctColor value (ARGB).
         */
        public void SetRgb(byte[] rgb)
        {
            // Correct it and save
            ctColor.SetRgb((rgb));
        }

        /**
         * Index into the clrScheme collection, referencing a particular sysClr or
         *  srgbClr value expressed in the Theme part.
         */
        public override int Theme
        {
            get
            {
                return ctColor.themeSpecified ? (int)ctColor.theme : (int)0;
            }
            set 
            {
                ctColor.theme = (uint)value;
            }
        }


        /**
         * Specifies the tint value applied to the ctColor.
         *
         * <p>
         * If tint is supplied, then it is applied to the RGB value of the ctColor to determine the final
         * ctColor applied.
         * </p>
         * <p>
         * The tint value is stored as a double from -1.0 .. 1.0, where -1.0 means 100% darken and
         * 1.0 means 100% lighten. Also, 0.0 means no Change.
         * </p>
         * <p>
         * In loading the RGB value, it is Converted to HLS where HLS values are (0..HLSMAX), where
         * HLSMAX is currently 255.
         * </p>
         * Here are some examples of how to apply tint to ctColor:
         * <blockquote>
         * <pre>
         * If (tint &lt; 0)
         * Lum' = Lum * (1.0 + tint)
         *
         * For example: Lum = 200; tint = -0.5; Darken 50%
         * Lum' = 200 * (0.5) =&gt; 100
         * For example: Lum = 200; tint = -1.0; Darken 100% (make black)
         * Lum' = 200 * (1.0-1.0) =&gt; 0
         * If (tint &gt; 0)
         * Lum' = Lum * (1.0-tint) + (HLSMAX - HLSMAX * (1.0-tint))
         * For example: Lum = 100; tint = 0.75; Lighten 75%
         *
         * Lum' = 100 * (1-.75) + (HLSMAX - HLSMAX*(1-.75))
         * = 100 * .25 + (255 - 255 * .25)
         * = 25 + (255 - 63) = 25 + 192 = 217
         * For example: Lum = 100; tint = 1.0; Lighten 100% (make white)
         * Lum' = 100 * (1-1) + (HLSMAX - HLSMAX*(1-1))
         * = 100 * 0 + (255 - 255 * 0)
         * = 0 + (255 - 0) = 255
         * </pre>
         * </blockquote>
         *
         * @return the tint value
         */
        public override double Tint
        {
            get
            {
                return ctColor.tint;
            }
            set 
            {
                ctColor.tint = value;
                ctColor.tintSpecified = true;
            }
        }

        /**
         * Returns the underlying XML bean
         *
         * @return the underlying XML bean
         */

        internal CT_Color GetCTColor()
        {
            return ctColor;
        }
        /// <summary>
        /// Checked type cast <tt>color</tt> to an XSSFColor.
        /// </summary>
        /// <param name="color">the color to type cast</param>
        /// <returns>the type casted color</returns>
        /// <exception cref="ArgumentException">if color is null or is not an instance of XSSFColor</exception>
        public static XSSFColor ToXSSFColor(IColor color)
        {
            // FIXME: this method would be more useful if it could convert any Color to an XSSFColor
            // Currently the only benefit of this method is to throw an IllegalArgumentException
            // instead of a ClassCastException.
            if (color != null && !(color is XSSFColor)) {
                throw new ArgumentException("Only XSSFColor objects are supported");
            }
            return (XSSFColor)color;
        }
        // Helper methods for {@link #equals(Object)}
        private bool SameIndexed(XSSFColor other)
        {
            if (IsIndexed == other.IsIndexed)
            {
                if (IsIndexed)
                {
                    return Indexed == other.Indexed;
                }
                return true;
            }
            return false;
        }
        private bool SameARGB(XSSFColor other)
        {
            if (IsRGB == other.IsRGB)
            {
                if (IsRGB)
                {
                    return Arrays.Equals(ARGB, other.ARGB);
                }
                return true;
            }
            return false;
        }
        private bool SameTheme(XSSFColor other)
        {
            if (IsThemed == other.IsThemed)
            {
                if (IsThemed)
                {
                    return Theme == other.Theme;
                }
                return true;
            }
            return false;
        }
        private bool SameTint(XSSFColor other)
        {
            if (HasTint == other.HasTint)
            {
                if (HasTint)
                {
                    return Tint == other.Tint;
                }
                return true;
            }
            return false;
        }
        private bool SameAuto(XSSFColor other)
        {
            return IsAuto == other.IsAuto;
        }
        public override int GetHashCode()
        {
            return ctColor.ToString().GetHashCode();
        }

        public override bool Equals(Object o)
        {
            if (o == null || !(o is XSSFColor))
                return false;

            XSSFColor other = (XSSFColor)o;

            // Compare each field in ctColor.
            // Cannot compare ctColor's XML string representation because equivalent
            // colors may have different relation namespace URI's
            return SameARGB(other)
                    && SameTheme(other)
                    && SameIndexed(other)
                    && SameTint(other)
                    && SameAuto(other);
        }
    }

}

