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
    using System.Text;
    using NPOI.HSSF.Util;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using SixLabors.ImageSharp.PixelFormats;

    /**
     * Represents a XSSF-style color (based on either a
     *  {@link NPOI.XSSF.UserModel.XSSFColor} or a
     *  {@link NPOI.HSSF.Record.Common.ExtendedColor} 
     */
    public abstract class ExtendedColor : IColor
    {
        protected void SetColor(Rgb24 clr)
        {
            RGB = (new byte[] { clr.R, clr.G, clr.B });
        }

        /**
         * A bool value indicating the color is automatic
         */
        public abstract bool IsAuto { get; set; }

        /**
         * A bool value indicating the color is indexed
         */
        public abstract bool IsIndexed { get; }

        /**
         * A bool value indicating the color is RGB / ARGB
         */
        public abstract bool IsRGB { get; }

        /**
         * A bool value indicating the color is from a Theme
         */
        public abstract bool IsThemed { get; }

        /**
         * Indexed Color value, if {@link #isIndexed()} is true
         */
        public abstract short Index { get; }

        public virtual short Indexed
        {
            get { return Index; }
        }

        /**
         * Index of Theme color, if {@link #isThemed()} is true
         */
        public abstract int Theme { get; set; }

        /**
         * Standard Red Green Blue ctColor value (RGB).
         * If there was an A (Alpha) value, it will be stripped.
         */
        public abstract byte[] RGB { get; set; }
        /**
         * Standard Alpha Red Green Blue ctColor value (ARGB).
         */
        public abstract byte[] ARGB { get; }

        /**
         * RGB or ARGB or null
         */
        protected abstract byte[] StoredRBG { get; }


        protected byte[] GetRGBOrARGB()
        {
            if (IsIndexed && Index > 0)
            {
                int indexNum = Index;
                var hashIndex = HSSFColor.GetIndexHash();
                HSSFColor indexed = null;
                if (hashIndex.ContainsKey(indexNum))
                    indexed = hashIndex[indexNum];
                if (indexed != null)
                {
                    byte[] rgb = new byte[3];
                    rgb[0] = (byte)indexed.GetTriplet()[0];
                    rgb[1] = (byte)indexed.GetTriplet()[1];
                    rgb[2] = (byte)indexed.GetTriplet()[2];
                    return rgb;
                }
            }

            // Grab the colour
            return StoredRBG;
        }

        /**
         * Standard Red Green Blue ctColor value (RGB) with applied tint.
         * Alpha values are ignored.
         */
        public byte[] RGBWithTint
        {
            get
            {
                byte[] rgb = StoredRBG;
                if (rgb != null)
                {
                    if (rgb.Length == 4)
                    {
                        byte[] tmp = new byte[3];
                        Array.Copy(rgb, 1, tmp, 0, 3);
                        rgb = tmp;
                    }
                    double tint = Tint;
                    for (int i = 0; i < rgb.Length; i++)
                    {
                        rgb[i] = ApplyTint(rgb[i] & 0xFF, tint);
                    }
                }
                return rgb;
            }
            
        }

        /**
         * Return the ARGB value in hex format, eg FF00FF00.
         * Works for both regular and indexed colours.
         */
        /**
        * Sets the ARGB value from hex format, eg FF0077FF.
        * Only works for regular (non-indexed) colours
        */
        public String ARGBHex
        {
            get
            {
                byte[] rgb = ARGB;
                if (rgb == null)
                {
                    return null;
                }
                StringBuilder sb = new StringBuilder();
                foreach (byte c in rgb)
                {
                    int i = c & 0xff;
                    String cs = string.Format("{0:x}", i);
                    if (cs.Length == 1)
                    {
                        sb.Append('0');
                    }
                    sb.Append(cs);
                }
                return sb.ToString().ToUpper();
            }
            set
            {
                if (value.Length == 6 || value.Length == 8)
                {
                    byte[] rgb = new byte[value.Length / 2];
                    for (int i = 0; i < rgb.Length; i++)
                    {
                        String part = value.Substring(i * 2, (i + 1) * 2-i*2);
                        rgb[i] = (byte)Int32.Parse(part, System.Globalization.NumberStyles.HexNumber);
                    }
                    RGB = (rgb);
                }
                else
                {
                    throw new ArgumentException("Must be of the form 112233 or FFEEDDCC");
                }
            }
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
        public abstract double Tint { get; set; }

    }

}