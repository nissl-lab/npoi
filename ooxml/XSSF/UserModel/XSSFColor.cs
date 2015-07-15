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
namespace NPOI.XSSF.UserModel
{

    /**
     * Represents a color in SpreadsheetML
     */
    public class XSSFColor : IColor
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

        public XSSFColor(System.Drawing.Color clr)
            : this()
        {

            ctColor.SetRgb((byte)clr.R, (byte)clr.G, (byte)clr.B);
        }

        public XSSFColor(byte[] rgb)
            : this()
        {

            ctColor.SetRgb(rgb);
        }

        /// <summary>
        ///A bool value indicating the ctColor is automatic and system ctColor dependent.
        /// </summary>
        public bool IsAuto
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


        /**
         * Indexed ctColor value. Only used for backwards compatibility. References a ctColor in indexedColors.
         */
        public short Indexed
        {
            get
            {
                return ctColor.indexedSpecified ? (short)ctColor.indexed : (short)0;
            }
            set 
            {
                ctColor.indexed = (uint)value;
                ctColor.indexedSpecified = true;
            }
        }


        public byte[] GetRgb()
        { 
            return this.RGB;
        }
        /**
         * Standard Red Green Blue ctColor value (RGB).
         * If there was an A (Alpha) value, it will be stripped.
         */
        public byte[] RGB
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
        }

        /**
         * Standard Alpha Red Green Blue ctColor value (ARGB).
         */
        public byte[] GetARgb()
        {
            byte[] rgb = GetRGBOrARGB();
            if (rgb == null) return null;

            if (rgb.Length == 3)
            {
                // Pad with the default Alpha
                byte[] tmp = new byte[4];
                tmp[0] = 255;
                Array.Copy(rgb, 0, tmp, 1, 3);
                return tmp;
            }
            else
            {
                return rgb;
            }
        }

        private byte[] GetRGBOrARGB()
        {
            byte[] rgb = null;

            if (ctColor.indexedSpecified && ctColor.indexed > 0)
            {
                HSSFColor indexed = (HSSFColor)HSSFColor.GetIndexHash()[(int)ctColor.indexed];
                if (indexed != null)
                {
                    rgb = new byte[3];
                    rgb[0] = (byte)indexed.GetTriplet()[0];
                    rgb[1] = (byte)indexed.GetTriplet()[1];
                    rgb[2] = (byte)indexed.GetTriplet()[2];
                    return rgb;
                }
            }

            if (!ctColor.IsSetRgb())
            {
                // No colour is available, sorry
                return null;
            }

            // Grab the colour
            rgb = ctColor.GetRgb();

            // Correct it as needed, and return
            return (rgb);
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

        /**
         * Return the ARGB value in hex format, eg FF00FF00.
         * Works for both regular and indexed colours. 
         */
        public String GetARGBHex()
        {
            StringBuilder sb = new StringBuilder();
            byte[] rgb = GetARgb();
            if (rgb == null)
            {
                return null;
            }
            foreach (byte c in rgb)
            {
                int i = (int)c;
                if (i < 0)
                {
                    i += 256;
                }
                String cs = StringUtil.ToHexString(i);
                if (cs.Length == 1)
                {
                    sb.Append('0');
                }
                sb.Append(cs);
            }
            return sb.ToString().ToUpper();
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
         * Index into the <clrScheme> collection, referencing a particular <sysClr> or
         *  <srgbClr> value expressed in the Theme part.
         */
        public int Theme
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
        public double Tint
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
         * @param tint the tint value
         */

        /**
         * Returns the underlying XML bean
         *
         * @return the underlying XML bean
         */

        internal CT_Color GetCTColor()
        {
            return ctColor;
        }

        public override int GetHashCode()
        {
            return ctColor.ToString().GetHashCode();
        }

        public override bool Equals(Object o)
        {
            if (o == null || !(o is XSSFColor)) return false;

            XSSFColor cf = (XSSFColor)o;
            return ctColor.ToString().Equals(cf.GetCTColor().ToString());
        }
    }

}

