/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License");you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.Util
{
    using System;
    using System.Collections;
    using System.IO;
    using System.Reflection;
    /**
     * Intends to provide support for the very evil index to triplet Issue and
     * will likely replace the color constants interface for HSSF 2.0.
     * This class Contains static inner class members for representing colors.
     * Each color has an index (for the standard palette in Excel (tm) ),
     * native (RGB) triplet and string triplet.  The string triplet Is as the
     * color would be represented by Gnumeric.  Having (string) this here Is a bit of a
     * collusion of function between HSSF and the HSSFSerializer but I think its
     * a reasonable one in this case.
     *
     * @author  Andrew C. Oliver (acoliver at apache dot org)
     * @author  Brian Sanders (bsanders at risklabs dot com) - full default color palette
     */
    public class HSSFColor : NPOI.SS.UserModel.IColor
    {
        private static Hashtable indexHash;
        public const short COLOR_NORMAL = 0x7fff;

        // TODO make subclass instances immutable
        /** Creates a new instance of HSSFColor */
        public HSSFColor()
        {
        }

        /**
         * this function returns all colors in a hastable.  Its not implemented as a
         * static member/staticly initialized because that would be dirty in a
         * server environment as it Is intended.  This means you'll eat the time
         * it takes to Create it once per request but you will not hold onto it
         * if you have none of those requests.
         *
         * @return a hashtable containing all colors keyed by <c>int</c> excel-style palette indexes
         */
        public static Hashtable GetIndexHash()
        {
            if (indexHash == null)
            {
                indexHash = CreateColorsByIndexMap();
            }
            return indexHash;
        }

        /**
         * This function returns all the Colours, stored in a Hashtable that
         *  can be edited. No caching is performed. If you don't need to edit
         *  the table, then call {@link #getIndexHash()} which returns a
         *  statically cached imuatable map of colours.
         */
        public static Hashtable GetMutableIndexHash()
        {
            return CreateColorsByIndexMap();
        }

        private static Hashtable CreateColorsByIndexMap()
        {
            HSSFColor[] colors = GetAllColors();
            Hashtable result = new Hashtable(colors.Length * 3 / 2);

            for (int i = 0; i < colors.Length; i++)
            {
                HSSFColor color = colors[i];

                int index1 = color.Indexed;
                if (result.ContainsKey(index1))
                {
                    HSSFColor prevColor = (HSSFColor)result[index1];
                    throw new InvalidDataException("Dup color index (" + index1
                            + ") for colors (" + prevColor.GetType().Name
                            + "),(" + color.GetType().Name + ")");
                }
                result[index1] = color;
            }

            for (int i = 0; i < colors.Length; i++)
            {
                HSSFColor color = colors[i];
                int index2 = GetIndex2(color);
                if (index2 == -1)
                {
                    // most colors don't have a second index
                    continue;
                }
                //if (result.ContainsKey(index2))
                //{
                //if (false)
                //{ // Many of the second indexes clash
                //    HSSFColor prevColor = (HSSFColor)result[index2];
                //    throw new InvalidDataException("Dup color index (" + index2
                //            + ") for colors (" + prevColor.GetType().Name
                //            + "),(" + color.GetType().Name + ")");
                //}
                //}
                result[index2] = color;
            }
            return result;
        }

        private static int GetIndex2(HSSFColor color)
        {
            FieldInfo f = color.GetType().GetField("Index2", BindingFlags.Static | BindingFlags.Public);
            if (f == null)
            {
                return -1;
            }

            short s = (short)f.GetValue(color);
            return Convert.ToInt32(s);
        }

        internal static HSSFColor[] GetAllColors()
        {
            return new HSSFColor[] {
                new Black(), new Brown(), new OliveGreen(), new DarkGreen(),
                new DarkTeal(), new DarkBlue(), new Indigo(), new Grey80Percent(),
                new Orange(), new DarkYellow(), new Green(), new Teal(), new Blue(),
                new BlueGrey(), new Grey50Percent(), new Red(), new LightOrange(), new Lime(),
                new SeaGreen(), new Aqua(), new LightBlue(), new Violet(), new Grey40Percent(),
                new Pink(), new Gold(), new Yellow(), new BrightGreen(), new Turquoise(),
                new DarkRed(), new SkyBlue(), new Plum(), new Grey25Percent(), new Rose(),
                new LightYellow(), new LightGreen(), new LightTurquoise(), new PaleBlue(),
                new Lavender(), new White(), new CornflowerBlue(), new LemonChiffon(),
                new Maroon(), new Orchid(), new Coral(), new RoyalBlue(),
                new LightCornflowerBlue(), new Tan(),
            };
        }

        /// <summary>
        /// this function returns all colors in a hastable.  Its not implemented as a
        /// static member/staticly initialized because that would be dirty in a
        /// server environment as it Is intended.  This means you'll eat the time
        /// it takes to Create it once per request but you will not hold onto it
        /// if you have none of those requests.
        /// </summary>
        /// <returns>a hashtable containing all colors keyed by String gnumeric-like triplets</returns>
        public static Hashtable GetTripletHash()
        {
            return CreateColorsByHexStringMap();
        }

        private static Hashtable CreateColorsByHexStringMap()
        {
            HSSFColor[] colors = GetAllColors();
            Hashtable result = new Hashtable(colors.Length * 3 / 2);

            for (int i = 0; i < colors.Length; i++)
            {
                HSSFColor color = colors[i];

                String hexString = color.GetHexString();
                if (result.ContainsKey(hexString))
                {
                    throw new InvalidDataException("Dup color hexString (" + hexString
                            + ") for color (" + color.GetType().Name + ")");
                }
                result[hexString] = color;
            }
            return result;
        }

        /**
         * @return index to the standard palette
         */

        public virtual short Indexed
        {
            get
            {
                return Black.Index;
            }
        }

        public byte[] RGB 
        {
            get { return this.GetTriplet(); }
        }
        /**
         * @return  triplet representation like that in Excel
         */
        public virtual byte[] GetTriplet()
        {
            return Black.Triplet;
        }

        // its a hack but its a good hack

        /**
         * @return a hex string exactly like a gnumeric triplet
         */

        public virtual String GetHexString()
        {
            return Black.HexString;
        }

        /**
         * Class BLACK
         *
         */

        public class Black : HSSFColor
        {
            public const short Index = 0x8;
            public static readonly byte[] Triplet = { 0, 0, 0 };
            public const string HexString = "0:0:0";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class BROWN
         *
         */

        public class Brown : HSSFColor
        {
            public const short Index = 0x3c;
            public static readonly byte[] Triplet = { 153, 51, 0 };
            public const string HexString = "9999:3333:0";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class OLIVE_GREEN
         *
         */

        public class OliveGreen: HSSFColor
        {
            public const short Index = 0x3b;
            public static readonly byte[] Triplet = { 51, 51, 0 };
            public const string HexString = "3333:3333:0";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class DARK_GREEN
         *
         */

        public class DarkGreen: HSSFColor
        {
            public const short Index = 0x3a;
            public static readonly byte[] Triplet = { 0, 51, 0 };
            public const string HexString = "0:3333:0";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class DARK_TEAL
         *
         */

        public class DarkTeal: HSSFColor
        {
            public const short Index = 0x38;
            public static readonly byte[] Triplet = { 0, 51, 102 };
            public const string HexString = "0:3333:6666";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class DARK_BLUE
         *
         */

        public class DarkBlue: HSSFColor
        {
            public const short Index = 0x12;
            public const short Index2 = 0x20;
            public static readonly byte[] Triplet = { 0, 0, 128 };
            public const string HexString = "0:0:8080";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class INDIGO
         *
         */

        public class Indigo: HSSFColor
        {
            public const short Index = 0x3e;
            public static readonly byte[] Triplet = { 51, 51, 153 };
            public const string HexString = "3333:3333:9999";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class GREY_80_PERCENT
         *
         */

        public class Grey80Percent: HSSFColor
        {
            public const short Index = 0x3f;
            public static readonly byte[] Triplet = { 51, 51, 51 };
            public const string HexString = "3333:3333:3333";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class DARK_RED
         *
         */

        public class DarkRed: HSSFColor
        {
            public const short Index = 0x10;
            public const short Index2 = 0x25;
            public static readonly byte[] Triplet = { 128, 0, 0 };
            public const string HexString = "8080:0:0";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class ORANGE
         *
         */

        public class Orange: HSSFColor
        {
            public const short Index = 0x35;
            public static readonly byte[] Triplet = { 255, 102, 0 };
            public const string HexString = "FFFF:6666:0";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class DARK_YELLOW
         *
         */

        public class DarkYellow: HSSFColor
        {
            public const short Index = 0x13;
            public static readonly byte[] Triplet = { 128, 128, 0 };
            public const string HexString = "8080:8080:0";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class GREEN
         *
         */

        public class Green: HSSFColor
        {
            public const short Index = 0x11;
            public static readonly byte[] Triplet = { 0, 128, 0 };
            public const string HexString = "0:8080:0";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class TEAL
         *
         */

        public class Teal: HSSFColor
        {
            public const short Index = 0x15;
            public const short Index2 = 0x26;
            public static readonly byte[] Triplet = { 0, 128, 128 };
            public const string HexString = "0:8080:8080";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class BLUE
         *
         */

        public class Blue: HSSFColor
        {
            public const short Index = 0xc;
            public const short Index2 = 0x27;
            public static readonly byte[] Triplet = { 0, 0, 255 };
            public const string HexString = "0:0:FFFF";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class BLUE_GREY
         *
         */

        public class BlueGrey: HSSFColor
        {
            public const short Index = 0x36;
            public static readonly byte[] Triplet = { 102, 102, 153 };
            public const string HexString = "6666:6666:9999";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class GREY_50_PERCENT
         *
         */

        public class Grey50Percent: HSSFColor
        {
            public const short Index = 0x17;
            public static readonly byte[] Triplet = { 128, 128, 128 };
            public const string HexString = "8080:8080:8080";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class RED
         *
         */

        public class Red: HSSFColor
        {
            public const short Index = 0xa;
            public static readonly byte[] Triplet = { 255, 0, 0 };
            public const string HexString = "FFFF:0:0";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class LIGHT_ORANGE
         *
         */

        public class LightOrange: HSSFColor
        {
            public const short Index = 0x34;
            public static readonly byte[] Triplet = { 255, 153, 0 };
            public const string HexString = "FFFF:9999:0";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class LIME
         *
         */
        public class Lime: HSSFColor
        {
            public const short Index = 0x32;
            public static readonly byte[] Triplet = { 153, 204, 0 };
            public const string HexString = "9999:CCCC:0";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class SEA_GREEN
         *
         */

        public class SeaGreen: HSSFColor
        {
            public const short Index = 0x39;
            public static readonly byte[] Triplet = { 51, 153, 102 };
            public const string HexString = "3333:9999:6666";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class AQUA
         *
         */

        public class Aqua: HSSFColor
        {
            public const short Index = 0x31;
            public static readonly byte[] Triplet = { 51, 204, 204 };
            public const string HexString = "3333:CCCC:CCCC";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        public class LightBlue: HSSFColor
        {
            public const short Index = 0x30;
            public static readonly byte[] Triplet = { 51, 102, 255 };
            public const string HexString = "3333:6666:FFFF";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        public class Violet: HSSFColor
        {
            public const short Index = 0x14;
            public const short Index2 = 0x24;
            public static readonly byte[] Triplet = { 128, 0, 128 };
            public const string HexString = "8080:0:8080";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class GREY_40_PERCENT
         *
         */

        public class Grey40Percent : HSSFColor
        {
            public const short Index = 0x37;
            public static readonly byte[] Triplet = { 150, 150, 150 };
            public const string HexString = "9696:9696:9696";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        public class Pink: HSSFColor
        {
            public const short Index = 0xe;
            public const short Index2 = 0x21;
            public static readonly byte[] Triplet = { 255, 0, 255 };
            public const string HexString = "FFFF:0:FFFF";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        public class Gold: HSSFColor
        {
            public const short Index = 0x33;
            public static readonly byte[] Triplet = { 255, 204, 0 };
            public const string HexString = "FFFF:CCCC:0";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }


        public class Yellow: HSSFColor
        {
            public const short Index = 0xd;
            public const short Index2 = 0x22;
            public static readonly byte[] Triplet = { 255, 255, 0 };
            public const string HexString = "FFFF:FFFF:0";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        public class BrightGreen: HSSFColor
        {
            public const short Index = 0xb;
            public const short Index2 = 0x23;
            public static readonly byte[] Triplet = { 0, 255, 0 };
            public const string HexString = "0:FFFF:0";

            public override String GetHexString()
            {
                return HexString;
            }

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }
        }

        /**
         * Class TURQUOISE
         *
         */

        public class Turquoise: HSSFColor
        {
            public const short Index = 0xf;
            public const short Index2 = 0x23;
            public static readonly byte[] Triplet = { 0, 255, 255 };
            public const string HexString = "0:FFFF:FFFF";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class SKY_BLUE
         *
         */

        public class SkyBlue: HSSFColor
        {
            public const short Index = 0x28;
            public static readonly byte[] Triplet = { 0, 204, 255 };
            public const string HexString = "0:CCCC:FFFF";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class PLUM
         *
         */

        public class Plum: HSSFColor
        {
            public const short Index = 0x3d;
            public const short Index2 = 0x19;
            public static readonly byte[] Triplet = { 153, 51, 102 };
            public const string HexString = "9999:3333:6666";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class GREY_25_PERCENT
         *
         */

        public class Grey25Percent: HSSFColor
        {
            public const short Index = 0x16;
            public static readonly byte[] Triplet = { 192, 192, 192 };
            public const string HexString = "C0C0:C0C0:C0C0";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class ROSE
         *
         */

        public class Rose: HSSFColor
        {
            public const short Index = 0x2d;
            public static readonly byte[] Triplet = { 255, 153, 204 };
            public const string HexString = "FFFF:9999:CCCC";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class TAN
         *
         */

        public class Tan: HSSFColor
        {
            public const short Index = 0x2f;
            public static readonly byte[] Triplet = { 255, 204, 153 };
            public const string HexString = "FFFF:CCCC:9999";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class LIGHT_YELLOW
         *
         */

        public class LightYellow: HSSFColor
        {
            public const short Index = 0x2b;
            public static readonly byte[] Triplet = { 255, 255, 153 };
            public const string HexString = "FFFF:FFFF:9999";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class LIGHT_GREEN
         *
         */

        public class LightGreen: HSSFColor
        {
            public const short Index = 0x2a;
            public static readonly byte[] Triplet = { 204, 255, 204 };
            public const string HexString = "CCCC:FFFF:CCCC";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class LIGHT_TURQUOISE
         *
         */

        public class LightTurquoise: HSSFColor
        {
            public const short Index = 0x29;
            public const short Index2 = 0x1b;
            public static readonly byte[] Triplet = { 204, 255, 255 };
            public const string HexString = "CCCC:FFFF:FFFF";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class PALE_BLUE
         *
         */

        public class PaleBlue: HSSFColor
        {
            public const short Index = 0x2c;
            public static readonly byte[] Triplet = { 153, 204, 255 };
            public const string HexString = "9999:CCCC:FFFF";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class LAVENDER
         *
         */

        public class Lavender: HSSFColor
        {
            public const short Index = 0x2e;
            public static readonly byte[] Triplet = { 204, 153, 255 };
            public const string HexString = "CCCC:9999:FFFF";


            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class WHITE
         *
         */

        public class White: HSSFColor
        {
            public const short Index = 0x9;
            public static readonly byte[] Triplet = { 255, 255, 255 };
            public const string HexString = "FFFF:FFFF:FFFF";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class CORNFLOWER_BLUE
         */
        public class CornflowerBlue: HSSFColor
        {
            public const short Index = 0x18;
            public static readonly byte[] Triplet = { 153, 153, 255 };
            public const string HexString = "9999:9999:FFFF";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }


        /**
         * Class LEMON_CHIFFON
         */
        public class LemonChiffon: HSSFColor
        {
            public const short Index = 0x1a;
            public static readonly byte[] Triplet = { 255, 255, 204 };

            public const string HexString = "FFFF:FFFF:CCCC";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class MAROON
         */
        public class Maroon: HSSFColor
        {
            public const short Index = 0x19;
            public static readonly byte[] Triplet = { 127, 0, 0 };
            public const string HexString = "8000:0:0";
            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class ORCHID
         */
        public class Orchid: HSSFColor
        {
            public const short Index = 0x1c;
            public static readonly byte[] Triplet = { 102, 0, 102 };
            public const string HexString = "6666:0:6666";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class CORAL
         */
        public class Coral : HSSFColor
        {
            public const short Index = 0x1d;
            public static readonly byte[] Triplet = { 255, 128, 128 };
            public const string HexString = "FFFF:8080:8080";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class ROYAL_BLUE
         */
        public class RoyalBlue : HSSFColor
        {
            public const short Index = 0x1e;
            public static readonly byte[] Triplet = { 0, 102, 204 };
            public const string HexString = "0:6666:CCCC";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Class LIGHT_CORNFLOWER_BLUE
         */
        public class LightCornflowerBlue : HSSFColor
        {
            public const short Index = 0x1f;
            public static readonly byte[] Triplet = { 204, 204, 255 };
            public const string HexString = "CCCC:CCCC:FFFF";

            public override short Indexed
            {
                get{return Index;}
            }

            public override byte[] GetTriplet()
            {
                return Triplet;
            }

            public override String GetHexString()
            {
                return HexString;
            }
        }

        /**
         * Special Default/Normal/Automatic color.
         * <i>Note:</i> This class Is NOT in the default HashTables returned by HSSFColor.
         * The index Is a special case which Is interpreted in the various SetXXXColor calls.
         *
         * @author Jason
         *
         */
        public class Automatic : HSSFColor
        {
            private static HSSFColor instance = new Automatic();

            public const short Index = 0x40;

            public override byte[] GetTriplet()
            {
                return Black.Triplet;
            }

            public override String GetHexString()
            {
                return Black.HexString;
            }

            public static HSSFColor GetInstance()
            {
                return instance;
            }
        }
    }
}
