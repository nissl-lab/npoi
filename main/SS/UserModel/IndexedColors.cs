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

        public static IndexedColors BLACK = new IndexedColors(8);
        public static IndexedColors WHITE = new IndexedColors(9);
        public static IndexedColors RED = new IndexedColors(10);
        public static IndexedColors BRIGHT_GREEN = new IndexedColors(11);
        public static IndexedColors BLUE = new IndexedColors(12);
        public static IndexedColors YELLOW = new IndexedColors(13);
        public static IndexedColors PINK = new IndexedColors(14);
        public static IndexedColors TURQUOISE = new IndexedColors(15);
        public static IndexedColors DARK_RED = new IndexedColors(16);
        public static IndexedColors GREEN = new IndexedColors(17);
        public static IndexedColors DARK_BLUE = new IndexedColors(18);
        public static IndexedColors DARK_YELLOW = new IndexedColors(19);
        public static IndexedColors VIOLET = new IndexedColors(20);
        public static IndexedColors TEAL = new IndexedColors(21);
        public static IndexedColors GREY_25_PERCENT = new IndexedColors(22);
        public static IndexedColors GREY_50_PERCENT = new IndexedColors(23);
        public static IndexedColors CORNFLOWER_BLUE = new IndexedColors(24);
        public static IndexedColors MAROON = new IndexedColors(25);
        public static IndexedColors LEMON_CHIFFON = new IndexedColors(26);
        public static IndexedColors ORCHID = new IndexedColors(28);
        public static IndexedColors CORAL = new IndexedColors(29);
        public static IndexedColors ROYAL_BLUE = new IndexedColors(30);
        public static IndexedColors LIGHT_CORNFLOWER_BLUE = new IndexedColors(31);
        public static IndexedColors SKY_BLUE = new IndexedColors(40);
        public static IndexedColors LIGHT_TURQUOISE = new IndexedColors(41);
        public static IndexedColors LIGHT_GREEN = new IndexedColors(42);
        public static IndexedColors LIGHT_YELLOW = new IndexedColors(43);
        public static IndexedColors PALE_BLUE = new IndexedColors(44);
        public static IndexedColors ROSE = new IndexedColors(45);
        public static IndexedColors LAVENDER = new IndexedColors(46);
        public static IndexedColors TAN = new IndexedColors(47);
        public static IndexedColors LIGHT_BLUE = new IndexedColors(48);
        public static IndexedColors AQUA = new IndexedColors(49);
        public static IndexedColors LIME = new IndexedColors(50);
        public static IndexedColors GOLD = new IndexedColors(51);
        public static IndexedColors LIGHT_ORANGE = new IndexedColors(52);
        public static IndexedColors ORANGE = new IndexedColors(53);
        public static IndexedColors BLUE_GREY = new IndexedColors(54);
        public static IndexedColors GREY_40_PERCENT = new IndexedColors(55);
        public static IndexedColors DARK_TEAL = new IndexedColors(56);
        public static IndexedColors SEA_GREEN = new IndexedColors(57);
        public static IndexedColors DARK_GREEN = new IndexedColors(58);
        public static IndexedColors OLIVE_GREEN = new IndexedColors(59);
        public static IndexedColors BROWN = new IndexedColors(60);
        public static IndexedColors PLUM = new IndexedColors(61);
        public static IndexedColors INDIGO = new IndexedColors(62);
        public static IndexedColors GREY_80_PERCENT = new IndexedColors(63);
        public static IndexedColors AUTOMATIC = new IndexedColors(64);

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
