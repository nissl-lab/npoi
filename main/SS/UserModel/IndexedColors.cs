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

using NPOI.HSSF.Util;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
namespace NPOI.SS.UserModel
{

    /**
     * A deprecated indexing scheme for colours that is still required for some records, and for backwards
     *  compatibility with OLE2 formats.
     *
     * <p>
     * Each element corresponds to a color index (zero-based). When using the default indexed color palette,
     * the values are not written out, but instead are implied. When the color palette has been modified from default,
     * then the entire color palette is used.
     * </p>
     *
     * @author Yegor Kozlov
     */
    public class IndexedColors
    {
        public static readonly IndexedColors Black;
        public static readonly IndexedColors White;
        public static readonly IndexedColors Red;
        public static readonly IndexedColors BrightGreen;
        public static readonly IndexedColors Blue;
        public static readonly IndexedColors Yellow;
        public static readonly IndexedColors Pink;
        public static readonly IndexedColors Turquoise;
        public static readonly IndexedColors DarkRed;
        public static readonly IndexedColors Green;
        public static readonly IndexedColors DarkBlue;
        public static readonly IndexedColors DarkYellow;
        public static readonly IndexedColors Violet;
        public static readonly IndexedColors Teal;
        public static readonly IndexedColors Grey25Percent;
        public static readonly IndexedColors Grey50Percent;
        public static readonly IndexedColors CornflowerBlue;
        public static readonly IndexedColors Maroon;
        public static readonly IndexedColors LemonChiffon;
        public static readonly IndexedColors Orchid;
        public static readonly IndexedColors Coral;
        public static readonly IndexedColors RoyalBlue;
        public static readonly IndexedColors LightCornflowerBlue;
        public static readonly IndexedColors SkyBlue;
        public static readonly IndexedColors LightTurquoise;
        public static readonly IndexedColors LightGreen;
        public static readonly IndexedColors LightYellow;
        public static readonly IndexedColors PaleBlue;
        public static readonly IndexedColors Rose;
        public static readonly IndexedColors Lavender;
        public static readonly IndexedColors Tan;
        public static readonly IndexedColors LightBlue;
        public static readonly IndexedColors Aqua;
        public static readonly IndexedColors Lime;
        public static readonly IndexedColors Gold;
        public static readonly IndexedColors LightOrange;
        public static readonly IndexedColors Orange;
        public static readonly IndexedColors BlueGrey;
        public static readonly IndexedColors Grey40Percent;
        public static readonly IndexedColors DarkTeal;
        public static readonly IndexedColors SeaGreen;
        public static readonly IndexedColors DarkGreen;
        public static readonly IndexedColors OliveGreen;
        public static readonly IndexedColors Brown ;
        public static readonly IndexedColors Plum;
        public static readonly IndexedColors Indigo;
        public static readonly IndexedColors Grey80Percent;
        public static readonly IndexedColors Automatic;

        private int index;
        private HSSFColor hssfColor;
        

        IndexedColors(int idx, HSSFColor color)
        {
            index = idx;
            this.hssfColor = color;
        }
        static Dictionary<string, IndexedColors> mappingName = null;
        static Dictionary<int, IndexedColors> mappingIndex = null;
        static IndexedColors()
        {
            Black = new IndexedColors(8, new HSSFColor.Black());
            White = new IndexedColors(9, new HSSFColor.White());
            Red = new IndexedColors(10, new HSSFColor.Red());
            BrightGreen = new IndexedColors(11, new HSSFColor.BrightGreen());
            Blue = new IndexedColors(12, new HSSFColor.Blue());
            Yellow = new IndexedColors(13, new HSSFColor.Yellow());
            Pink = new IndexedColors(14, new HSSFColor.Pink());
            Turquoise = new IndexedColors(15, new HSSFColor.Turquoise());
            DarkRed = new IndexedColors(16, new HSSFColor.DarkRed());
            Green = new IndexedColors(17, new HSSFColor.Green());
            DarkBlue = new IndexedColors(18, new HSSFColor.DarkBlue());
            DarkYellow = new IndexedColors(19, new HSSFColor.DarkYellow());
            Violet = new IndexedColors(20, new HSSFColor.Violet());
            Teal = new IndexedColors(21, new HSSFColor.Teal());
            Grey25Percent = new IndexedColors(22, new HSSFColor.Grey25Percent());
            Grey50Percent = new IndexedColors(23, new HSSFColor.Grey50Percent());
            CornflowerBlue = new IndexedColors(24, new HSSFColor.CornflowerBlue());
            Maroon = new IndexedColors(25, new HSSFColor.Maroon());
            LemonChiffon = new IndexedColors(26, new HSSFColor.LemonChiffon());
            Orchid = new IndexedColors(28, new HSSFColor.Orchid());
            Coral = new IndexedColors(29, new HSSFColor.Coral());
            RoyalBlue = new IndexedColors(30, new HSSFColor.RoyalBlue());
            LightCornflowerBlue = new IndexedColors(31,new HSSFColor.LightCornflowerBlue());
            SkyBlue = new IndexedColors(40, new HSSFColor.SkyBlue());
            LightTurquoise = new IndexedColors(41, new HSSFColor.LightTurquoise());
            LightGreen = new IndexedColors(42, new HSSFColor.LightGreen());
            LightYellow = new IndexedColors(43, new HSSFColor.LightYellow());
            PaleBlue = new IndexedColors(44, new HSSFColor.PaleBlue());
            Rose = new IndexedColors(45, new HSSFColor.Rose());
            Lavender = new IndexedColors(46, new HSSFColor.Lavender());
            Tan = new IndexedColors(47, new HSSFColor.Tan());
            LightBlue = new IndexedColors(48, new HSSFColor.LightBlue());
            Aqua = new IndexedColors(49, new HSSFColor.Aqua());
            Lime = new IndexedColors(50, new HSSFColor.Lime());
            Gold = new IndexedColors(51, new HSSFColor.Gold());
            LightOrange = new IndexedColors(52, new HSSFColor.LightOrange());
            Orange = new IndexedColors(53, new HSSFColor.Orange());
            BlueGrey = new IndexedColors(54, new HSSFColor.BlueGrey());
            Grey40Percent = new IndexedColors(55, new HSSFColor.Grey40Percent());
            DarkTeal = new IndexedColors(56, new HSSFColor.DarkTeal());
            SeaGreen = new IndexedColors(57, new HSSFColor.SeaGreen());
            DarkGreen = new IndexedColors(58, new HSSFColor.DarkGreen());
            OliveGreen = new IndexedColors(59, new HSSFColor.OliveGreen());
            Brown = new IndexedColors(60, new HSSFColor.Brown());
            Plum = new IndexedColors(61, new HSSFColor.Plum());
            Indigo = new IndexedColors(62, new HSSFColor.Indigo());
            Grey80Percent = new IndexedColors(63, new HSSFColor.Grey80Percent());
            Automatic = new IndexedColors(64, new HSSFColor.Automatic());


            mappingName = new Dictionary<string, IndexedColors>();
            mappingName.Add("black", IndexedColors.Black);
            mappingName.Add("white", IndexedColors.White);
            mappingName.Add("red", IndexedColors.Red);
            mappingName.Add("brightgreen", IndexedColors.BrightGreen);
            mappingName.Add("blue", IndexedColors.Blue);
            mappingName.Add("yellow", IndexedColors.Yellow);
            mappingName.Add("pink", IndexedColors.Pink);
            mappingName.Add("turquoise", IndexedColors.Turquoise);
            mappingName.Add("darkred", IndexedColors.DarkRed);
            mappingName.Add("green", IndexedColors.Green);
            mappingName.Add("darkblue", IndexedColors.DarkBlue);
            mappingName.Add("darkyellow", IndexedColors.DarkYellow);
            mappingName.Add("violet", IndexedColors.Violet);
            mappingName.Add("teal", IndexedColors.Teal);
            mappingName.Add("grey25percent", IndexedColors.Grey25Percent);
            mappingName.Add("grey50percent", IndexedColors.Grey50Percent);
            mappingName.Add("cornflowerblue", IndexedColors.CornflowerBlue);
            mappingName.Add("maroon", IndexedColors.Maroon);
            mappingName.Add("lemonchiffon", IndexedColors.LemonChiffon);
            mappingName.Add("orchid", IndexedColors.Orchid);
            mappingName.Add("coral", IndexedColors.Coral);
            mappingName.Add("royalblue", IndexedColors.RoyalBlue);
            mappingName.Add("lightcornflowerblue", IndexedColors.LightCornflowerBlue);
            mappingName.Add("skyblue", IndexedColors.SkyBlue);
            mappingName.Add("lightturquoise", IndexedColors.LightTurquoise);
            mappingName.Add("lightgreen", IndexedColors.LightGreen);
            mappingName.Add("lightyellow", IndexedColors.LightYellow);
            mappingName.Add("paleblue", IndexedColors.PaleBlue);
            mappingName.Add("rose", IndexedColors.Rose);
            mappingName.Add("lavender", IndexedColors.Lavender);
            mappingName.Add("tan", IndexedColors.Tan);
            mappingName.Add("lightblue", IndexedColors.LightBlue);
            mappingName.Add("aqua", IndexedColors.Aqua);
            mappingName.Add("lime", IndexedColors.Lime);
            mappingName.Add("gold", IndexedColors.Gold);
            mappingName.Add("lightorange", IndexedColors.LightOrange);
            mappingName.Add("orange", IndexedColors.Orange);
            mappingName.Add("bluegrey", IndexedColors.BlueGrey);
            mappingName.Add("grey40percent", IndexedColors.Grey40Percent);
            mappingName.Add("darkteal", IndexedColors.DarkTeal);
            mappingName.Add("seagreen", IndexedColors.SeaGreen);
            mappingName.Add("darkgreen", IndexedColors.DarkGreen);
            mappingName.Add("olivergreen", IndexedColors.OliveGreen);
            mappingName.Add("brown", IndexedColors.Brown);
            mappingName.Add("plum", IndexedColors.Plum);
            mappingName.Add("indigo", IndexedColors.Indigo);
            mappingName.Add("grey80percent", IndexedColors.Grey80Percent);
            mappingName.Add("automatic", IndexedColors.Automatic);


            mappingIndex = new Dictionary<int, IndexedColors>();
            mappingIndex.Add(8, IndexedColors.Black);
            mappingIndex.Add(9, IndexedColors.White);
            mappingIndex.Add(10, IndexedColors.Red);
            mappingIndex.Add(11, IndexedColors.BrightGreen);
            mappingIndex.Add(12, IndexedColors.Blue);
            mappingIndex.Add(13, IndexedColors.Yellow);
            mappingIndex.Add(14, IndexedColors.Pink);
            mappingIndex.Add(15, IndexedColors.Turquoise);
            mappingIndex.Add(16, IndexedColors.DarkRed);
            mappingIndex.Add(17, IndexedColors.Green);
            mappingIndex.Add(18, IndexedColors.DarkBlue);
            mappingIndex.Add(19, IndexedColors.DarkYellow);
            mappingIndex.Add(20, IndexedColors.Violet);
            mappingIndex.Add(21, IndexedColors.Teal);
            mappingIndex.Add(22, IndexedColors.Grey25Percent);
            mappingIndex.Add(23, IndexedColors.Grey50Percent);
            mappingIndex.Add(24, IndexedColors.CornflowerBlue);
            mappingIndex.Add(25, IndexedColors.Maroon);
            mappingIndex.Add(26, IndexedColors.LemonChiffon);
            mappingIndex.Add(28, IndexedColors.Orchid);
            mappingIndex.Add(29, IndexedColors.Coral);
            mappingIndex.Add(30, IndexedColors.RoyalBlue);
            mappingIndex.Add(31, IndexedColors.LightCornflowerBlue);
            mappingIndex.Add(40, IndexedColors.SkyBlue);
            mappingIndex.Add(41, IndexedColors.LightTurquoise);
            mappingIndex.Add(42, IndexedColors.LightGreen);
            mappingIndex.Add(43, IndexedColors.LightYellow);
            mappingIndex.Add(44, IndexedColors.PaleBlue);
            mappingIndex.Add(45, IndexedColors.Rose);
            mappingIndex.Add(46, IndexedColors.Lavender);
            mappingIndex.Add(47, IndexedColors.Tan);
            mappingIndex.Add(48, IndexedColors.LightBlue);
            mappingIndex.Add(49, IndexedColors.Aqua);
            mappingIndex.Add(50, IndexedColors.Lime);
            mappingIndex.Add(51, IndexedColors.Gold);
            mappingIndex.Add(52, IndexedColors.LightOrange);
            mappingIndex.Add(53, IndexedColors.Orange);
            mappingIndex.Add(54, IndexedColors.BlueGrey);
            mappingIndex.Add(55, IndexedColors.Grey40Percent);
            mappingIndex.Add(56, IndexedColors.DarkTeal);
            mappingIndex.Add(57, IndexedColors.SeaGreen);
            mappingIndex.Add(58, IndexedColors.DarkGreen);
            mappingIndex.Add(59, IndexedColors.OliveGreen);
            mappingIndex.Add(60, IndexedColors.Brown);
            mappingIndex.Add(61, IndexedColors.Plum);
            mappingIndex.Add(62, IndexedColors.Indigo);
            mappingIndex.Add(63, IndexedColors.Grey80Percent);
            mappingIndex.Add(64, IndexedColors.Automatic);
        }
        public static IndexedColors ValueOf(string colorName)
        {
            if (mappingName.ContainsKey(colorName.ToLower()))
                return mappingName[colorName.ToLower()];

            return null;
        }
        public static IndexedColors ValueOf(int index)
        {
            if(mappingIndex.ContainsKey(index))
                return mappingIndex[index];
            throw new ArgumentException("Illegal IndexedColor index: " + index);
        }

        /**
         * 
         *
         * @param index the index of the color
         * @return the corresponding IndexedColors enum
         * @throws IllegalArgumentException if index is not a valid IndexedColors
         * @since 3.15-beta2
         */
        public static IndexedColors FromInt(int index)
        {
            return ValueOf(index);
        }

        public byte[] RGB
        {
            get { return hssfColor.RGB; }
        }
        public string HexString
        {
            get
            {
                StringBuilder stringBuilder = new StringBuilder(7);
                stringBuilder.Append('#');

                byte[] rgb = this.hssfColor.RGB;
                foreach (byte s in rgb)
                {
                    stringBuilder.Append(s.ToString("x2"));
                }
                return stringBuilder.ToString();
            }
        }

        /**
         * Returns index of this color
         *
         * @return index of this color
         */
        public short Index
        {
            get
            {
                return (short)index;
            }
        }
    }
}
