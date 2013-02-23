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
        public static readonly IndexedColors Black = new IndexedColors(8);
        public static readonly IndexedColors White = new IndexedColors(9);
        public static readonly IndexedColors Red = new IndexedColors(10);
        public static readonly IndexedColors BrightGreen = new IndexedColors(11);
        public static readonly IndexedColors Blue = new IndexedColors(12);
        public static readonly IndexedColors Yellow = new IndexedColors(13);
        public static readonly IndexedColors Pink = new IndexedColors(14);
        public static readonly IndexedColors Turquoise = new IndexedColors(15);
        public static readonly IndexedColors DarkRed = new IndexedColors(16);
        public static readonly IndexedColors Green = new IndexedColors(17);
        public static readonly IndexedColors DarkBlue = new IndexedColors(18);
        public static readonly IndexedColors DarkYellow = new IndexedColors(19);
        public static readonly IndexedColors Violet = new IndexedColors(20);
        public static readonly IndexedColors Teal = new IndexedColors(21);
        public static readonly IndexedColors Grey25Percent = new IndexedColors(22);
        public static readonly IndexedColors Grey50Percent = new IndexedColors(23);
        public static readonly IndexedColors CornflowerBlue = new IndexedColors(24);
        public static readonly IndexedColors Maroon = new IndexedColors(25);
        public static readonly IndexedColors LemonChiffon = new IndexedColors(26);
        public static readonly IndexedColors Orchid = new IndexedColors(28);
        public static readonly IndexedColors Coral = new IndexedColors(29);
        public static readonly IndexedColors RoyalBlue = new IndexedColors(30);
        public static readonly IndexedColors LightCornflowerBlue = new IndexedColors(31);
        public static readonly IndexedColors SkyBlue = new IndexedColors(40);
        public static readonly IndexedColors LightTurquoise = new IndexedColors(41);
        public static readonly IndexedColors LightGreen = new IndexedColors(42);
        public static readonly IndexedColors LightYellow = new IndexedColors(43);
        public static readonly IndexedColors PaleBlue = new IndexedColors(44);
        public static readonly IndexedColors Rose = new IndexedColors(45);
        public static readonly IndexedColors Lavender = new IndexedColors(46);
        public static readonly IndexedColors Tan = new IndexedColors(47);
        public static readonly IndexedColors LightBlue = new IndexedColors(48);
        public static readonly IndexedColors Aqua = new IndexedColors(49);
        public static readonly IndexedColors Lime = new IndexedColors(50);
        public static readonly IndexedColors Gold = new IndexedColors(51);
        public static readonly IndexedColors LightOrange = new IndexedColors(52);
        public static readonly IndexedColors Orange = new IndexedColors(53);
        public static readonly IndexedColors BlueGrey = new IndexedColors(54);
        public static readonly IndexedColors Grey40Percent = new IndexedColors(55);
        public static readonly IndexedColors DarkTeal = new IndexedColors(56);
        public static readonly IndexedColors SeaGreen = new IndexedColors(57);
        public static readonly IndexedColors DarkGreen = new IndexedColors(58);
        public static readonly IndexedColors OliveGreen = new IndexedColors(59);
        public static readonly IndexedColors Brown = new IndexedColors(60);
        public static readonly IndexedColors Plum = new IndexedColors(61);
        public static readonly IndexedColors Indigo = new IndexedColors(62);
        public static readonly IndexedColors Grey80Percent = new IndexedColors(63);
        public static readonly IndexedColors Automatic = new IndexedColors(64);

        private int index;

        IndexedColors(int idx)
        {
            index = idx;
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
