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
    using System.Collections;
    using NPOI.Util.Collections;
    using System.Globalization;

    /// <summary>
    /// Stores width and height details about a font.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class FontDetails
    {
        private String fontName;
        private int height;
        private Hashtable charWidths = new Hashtable();

        /// <summary>
        /// Construct the font details with the given name and height.
        /// </summary>
        /// <param name="fontName">The font name.</param>
        /// <param name="height">The height of the font.</param>
        public FontDetails(String fontName, int height)
        {
            this.fontName = fontName;
            this.height = height;
        }

        /// <summary>
        /// Gets the name of the font.
        /// </summary>
        /// <returns></returns>
        public String GetFontName()
        {
            return fontName;
        }

        /// <summary>
        /// Gets the height.
        /// </summary>
        /// <returns></returns>
        public int GetHeight()
        {
            return height;
        }

        /// <summary>
        /// Adds the char.
        /// </summary>
        /// <param name="c">The c.</param>
        /// <param name="width">The width.</param>
        public void AddChar(char c, int width)
        {
            charWidths[c] = width;
        }

        /// <summary>
        /// Retrieves the width of the specified Char.  If the metrics for
        /// a particular Char are not available it defaults to returning the
        /// width for the 'W' Char.
        /// </summary>
        /// <param name="c">The character.</param>
        /// <returns></returns>
        public int GetCharWidth(char c)
        {
            object widthInteger = charWidths[c];
            if (widthInteger == null)
                return 'W' == c ? 0 : GetCharWidth('W');
            else
                return (int)widthInteger;
        }

        /// <summary>
        /// Adds the chars.
        /// </summary>
        /// <param name="Chars">The chars.</param>
        /// <param name="widths">The widths.</param>
        public void AddChars(char[] Chars, int[] widths)
        {
            for (int i = 0; i < Chars.Length; i++)
            {
                if (Chars[i] != ' ')
                {
                    charWidths[Chars[i]] = widths[i];
                }
            }
        }

        /// <summary>
        /// Builds the font height property.
        /// </summary>
        /// <param name="fontName">Name of the font.</param>
        /// <returns></returns>
        public static String BuildFontHeightProperty(String fontName)
        {
            return "font." + fontName + ".height";
        }
        /// <summary>
        /// Builds the font widths property.
        /// </summary>
        /// <param name="fontName">Name of the font.</param>
        /// <returns></returns>
        public static String BuildFontWidthsProperty(String fontName)
        {
            return "font." + fontName + ".widths";
        }
        /// <summary>
        /// Builds the font chars property.
        /// </summary>
        /// <param name="fontName">Name of the font.</param>
        /// <returns></returns>
        public static String BuildFontCharsProperty(String fontName)
        {
            return "font." + fontName + ".characters";
        }

        /// <summary>
        /// Create an instance of 
        /// <c>FontDetails</c>
        ///  by loading them from the
        /// provided property object.
        /// </summary>
        /// <param name="fontName">the font name.</param>
        /// <param name="fontMetricsProps">the property object holding the details of this
        /// particular font.</param>
        /// <returns>a new FontDetails instance.</returns>
        public static FontDetails Create(String fontName, Properties fontMetricsProps)
        {
            String heightStr = fontMetricsProps[BuildFontHeightProperty(fontName)];
            String widthsStr = fontMetricsProps[BuildFontWidthsProperty(fontName)];
            String CharsStr = fontMetricsProps[BuildFontCharsProperty(fontName)];

            // Ensure that this Is a font we know about
            if (heightStr == null || widthsStr == null || CharsStr == null)
            {
                // We don't know all we need to about this font
                // Since we don't know its sizes, we can't work with it
                throw new ArgumentException("The supplied FontMetrics doesn't know about the font '" + fontName + "', so we can't use it. Please Add it to your font metrics file (see StaticFontMetrics.GetFontDetails");
            }

            int height = int.Parse(heightStr, CultureInfo.InvariantCulture);
            FontDetails d = new FontDetails(fontName, height);
            String[] CharsStrArray = Split(CharsStr, ",", -1);
            String[] widthsStrArray = Split(widthsStr, ",", -1);
            if (CharsStrArray.Length != widthsStrArray.Length)
                throw new Exception("Number of Chars does not number of widths for font " + fontName);
            for (int i = 0; i < widthsStrArray.Length; i++)
            {
                if (CharsStrArray[i].Trim().Length != 0)
                    d.AddChar(CharsStrArray[i].Trim()[0], int.Parse(widthsStrArray[i], CultureInfo.InvariantCulture));
            }
            return d;
        }

        /// <summary>
        /// Gets the width of all Chars in a string.
        /// </summary>
        /// <param name="str">The string to measure.</param>
        /// <returns>The width of the string for a 10 point font.</returns>
        public int GetStringWidth(String str)
        {
            int width = 0;
            for (int i = 0; i < str.Length; i++)
            {
                width += GetCharWidth(str[i]);
            }
            return width;
        }

        /// <summary>
        /// Split the given string into an array of strings using the given
        /// delimiter.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="separator">The separator.</param>
        /// <param name="max">The max.</param>
        /// <returns></returns>
        private static String[] Split(String text, String separator, int max)
        {
            String[] list = text.Split(separator.ToCharArray());
            return list;
        }


    }
}