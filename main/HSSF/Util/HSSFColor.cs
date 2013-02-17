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

namespace NPOI.HSSF.Util
{

    using System;
    using System.IO;
    using System.Collections;
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

                int index1 = color.GetIndex();
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

            FieldInfo f;

            f = color.GetType().GetField("index2", BindingFlags.Static | BindingFlags.Public);
            if (f == null)
            {
                return -1;
            }

            short s;
            try
            {
                s = (short)f.GetValue(color);
            }
            catch (ArgumentException)
            {
                throw;
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            return Convert.ToInt32(s);
        }

        private static HSSFColor[] GetAllColors()
        {

            return new HSSFColor[] {
                new BLACK(), new BROWN(), new OLIVE_GREEN(), new DARK_GREEN(),
                new DARK_TEAL(), new DARK_BLUE(), new INDIGO(), new GREY_80_PERCENT(),
                new ORANGE(), new DARK_YELLOW(), new GREEN(), new TEAL(), new BLUE(),
                new BLUE_GREY(), new GREY_50_PERCENT(), new RED(), new LIGHT_ORANGE(), new LIME(),
                new SEA_GREEN(), new AQUA(), new LIGHT_BLUE(), new VIOLET(), new GREY_40_PERCENT(),
                new PINK(), new GOLD(), new YELLOW(), new BRIGHT_GREEN(), new TURQUOISE(),
                new DARK_RED(), new SKY_BLUE(), new PLUM(), new GREY_25_PERCENT(), new ROSE(),
                new LIGHT_YELLOW(), new LIGHT_GREEN(), new LIGHT_TURQUOISE(), new PALE_BLUE(),
                new LAVENDER(), new WHITE(), new CORNFLOWER_BLUE(), new LEMON_CHIFFON(),
                new MAROON(), new ORCHID(), new CORAL(), new ROYAL_BLUE(),
                new LIGHT_CORNFLOWER_BLUE(), new TAN(),
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

        public virtual short GetIndex()
        {
            return BLACK.index;
        }

        /**
         * @return  triplet representation like that in Excel
         */

        public virtual short[] GetTriplet()
        {
            return BLACK.triplet; 
        }

        // its a hack but its a good hack

        /**
         * @return a hex string exactly like a gnumeric triplet
         */

        public virtual String GetHexString()
        {
            return BLACK.hexString;
        }

        /**
         * Class BLACK
         *
         */

        public class BLACK : HSSFColor
        {
            public static short index = 0x8;
            public static short[] triplet =
            {
                0, 0, 0
            };
            public static String hexString = "0:0:0";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class BROWN
         *
         */

        public class BROWN : HSSFColor
        {
            public static short index = 0x3c;
            public static short[] triplet =
            {
                153, 51, 0
            };
            public static String hexString = "9999:3333:0";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }
            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class OLIVE_GREEN
         *
         */

        public class OLIVE_GREEN: HSSFColor
        {
            public static short index = 0x3b;
            public static short[] triplet =
            {
                51, 51, 0
            };
            public static String hexString = "3333:3333:0";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class DARK_GREEN
         *
         */

        public class DARK_GREEN: HSSFColor
        {
            public static short index = 0x3a;
            public static short[] triplet =
            {
                0, 51, 0
            };
            public static String hexString = "0:3333:0";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class DARK_TEAL
         *
         */

        public class DARK_TEAL: HSSFColor
        {
            public static short index = 0x38;
            public static short[] triplet =
            {
                0, 51, 102
            };
            public static String hexString = "0:3333:6666";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class DARK_BLUE
         *
         */

        public class DARK_BLUE: HSSFColor
        {
            public static short index = 0x12;
            public static short index2 = 0x20;
            public static short[] triplet =
            {
                0, 0, 128
            };
            public static String hexString = "0:0:8080";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class INDIGO
         *
         */

        public class INDIGO: HSSFColor
        {
            public static short index = 0x3e;
            public static short[] triplet =
            {
                51, 51, 153
            };
            public static String hexString = "3333:3333:9999";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class GREY_80_PERCENT
         *
         */

        public class GREY_80_PERCENT: HSSFColor
        {
            public static short index = 0x3f;
            public static short[] triplet =
            {
                51, 51, 51
            };
            public static String hexString = "3333:3333:3333";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class DARK_RED
         *
         */

        public class DARK_RED: HSSFColor
        {
            public static short index = 0x10;
            public static short index2 = 0x25;
            public static short[] triplet =
            {
                128, 0, 0
            };
            public static String hexString = "8080:0:0";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class ORANGE
         *
         */

        public class ORANGE: HSSFColor
        {
            public static short index = 0x35;
            public static short[] triplet =
            {
                255, 102, 0
            };
            public static String hexString = "FFFF:6666:0";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }
            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class DARK_YELLOW
         *
         */

        public class DARK_YELLOW: HSSFColor
        {
            public static short index = 0x13;
            public static short[] triplet =
            {
                128, 128, 0
            };
            public static String hexString = "8080:8080:0";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class GREEN
         *
         */

        public class GREEN: HSSFColor
        {
            public static short index = 0x11;
            public static short[] triplet =
            {
                0, 128, 0
            };
            public static String hexString = "0:8080:0";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class TEAL
         *
         */

        public class TEAL: HSSFColor
        {
            public static short index = 0x15;
            public static short index2 = 0x26;
            public static short[] triplet =
            {
                0, 128, 128
            };
            public static String hexString = "0:8080:8080";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class BLUE
         *
         */

        public class BLUE: HSSFColor
        {
            public static short index = 0xc;
            public static short index2 = 0x27;
            public static short[] triplet =
        {
            0, 0, 255
        };
            public static String hexString = "0:0:FFFF";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class BLUE_GREY
         *
         */

        public class BLUE_GREY: HSSFColor
        {
            public static short index = 0x36;
            public static short[] triplet =
            {
                102, 102, 153
            };
            public static String hexString = "6666:6666:9999";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class GREY_50_PERCENT
         *
         */

        public class GREY_50_PERCENT: HSSFColor
        {
            public static short index = 0x17;
            public static short[] triplet =
            {
                128, 128, 128
            };
            public static String hexString = "8080:8080:8080";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class RED
         *
         */

        public class RED: HSSFColor
        {
            public static short index = 0xa;
            public static short[] triplet =
            {
                255, 0, 0
            };
            public static String hexString = "FFFF:0:0";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class LIGHT_ORANGE
         *
         */

        public class LIGHT_ORANGE: HSSFColor
        {
            public static short index = 0x34;
            public static short[] triplet =
            {
                255, 153, 0
            };
            public static String hexString = "FFFF:9999:0";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class LIME
         *
         */
        public class LIME: HSSFColor
        {
            public static short index = 0x32;
            public static short[] triplet =
            {
                153, 204, 0
            };
            public static String hexString = "9999:CCCC:0";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class SEA_GREEN
         *
         */

        public class SEA_GREEN: HSSFColor
        {
            public static short index = 0x39;
            public static short[] triplet =
            {
                51, 153, 102
            };
            public static String hexString = "3333:9999:6666";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class AQUA
         *
         */

        public class AQUA: HSSFColor
        {
            public static short index = 0x31;
            public static short[] triplet =
            {
                51, 204, 204
            };
            public static String hexString = "3333:CCCC:CCCC";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        public class LIGHT_BLUE: HSSFColor
        {
            public static short index = 0x30;
            public static short[] triplet =
            {
                51, 102, 255
            };
            public static String hexString = "3333:6666:FFFF";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        public class VIOLET: HSSFColor
        {
            public static short index = 0x14;
            public static short index2 = 0x24;
            public static short[] triplet =
            {
                128, 0, 128
            };
            public static String hexString = "8080:0:8080";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class GREY_40_PERCENT
         *
         */

        public class GREY_40_PERCENT : HSSFColor
        {
            public static short index = 0x37;
            public static short[] triplet =
            {
                150, 150, 150
            };
            public static String hexString = "9696:9696:9696";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        public class PINK: HSSFColor
        {
            public static short index = 0xe;
            public static short index2 = 0x21;
            public static short[] triplet =
            {
                255, 0, 255
            };
            public static String hexString = "FFFF:0:FFFF";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        public class GOLD: HSSFColor
        {
            public static short index = 0x33;
            public static short[] triplet =
            {
                255, 204, 0
            };
            public static String hexString = "FFFF:CCCC:0";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }


        public class YELLOW: HSSFColor
        {
            public static short index = 0xd;
            public static short index2 = 0x22;
            public static short[] triplet =
            {
                255, 255, 0
            };
            public static String hexString = "FFFF:FFFF:0";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet; 
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        public class BRIGHT_GREEN: HSSFColor
        {
            public static short index = 0xb;
            public static short index2 = 0x23;
            public static short[] triplet =
            {
                0, 255, 0
            };
            public static String hexString = "0:FFFF:0";

            public override String GetHexString()
            {
                return hexString; 
            }

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }
        }

        /**
         * Class TURQUOISE
         *
         */

        public class TURQUOISE: HSSFColor
        {
            public static short index = 0xf;
            public static short index2 = 0x23;
            public static short[] triplet =
            {
                0, 255, 255
            };
            public static String hexString = "0:FFFF:FFFF";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class SKY_BLUE
         *
         */

        public class SKY_BLUE: HSSFColor
        {
            public static short index = 0x28;
            public static short[] triplet =
            {
                0, 204, 255
            };
            public static String hexString = "0:CCCC:FFFF";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class PLUM
         *
         */

        public class PLUM: HSSFColor
        {
            public static short index = 0x3d;
            public static short index2 = 0x19;
            public static short[] triplet =
            {
                153, 51, 102
            };
            public static String hexString = "9999:3333:6666";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class GREY_25_PERCENT
         *
         */

        public class GREY_25_PERCENT: HSSFColor
        {
            public static short index = 0x16;
            public static short[] triplet =
            {
                192, 192, 192
            };
            public static String hexString = "C0C0:C0C0:C0C0";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class ROSE
         *
         */

        public class ROSE: HSSFColor
        {
            public static short index = 0x2d;
            public static short[] triplet =
            {
                255, 153, 204
            };
            public static String hexString = "FFFF:9999:CCCC";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class TAN
         *
         */

        public class TAN: HSSFColor
        {
            public static short index = 0x2f;
            public static short[] triplet =
            {
                255, 204, 153
            };
            public static String hexString = "FFFF:CCCC:9999";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class LIGHT_YELLOW
         *
         */

        public class LIGHT_YELLOW: HSSFColor
        {
            public static short index = 0x2b;
            public static short[] triplet =
            {
                255, 255, 153
            };
            public static String hexString = "FFFF:FFFF:9999";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class LIGHT_GREEN
         *
         */

        public class LIGHT_GREEN: HSSFColor
        {
            public static short index = 0x2a;
            public static short[] triplet =
            {
                204, 255, 204
            };
            public static String hexString = "CCCC:FFFF:CCCC";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class LIGHT_TURQUOISE
         *
         */

        public class LIGHT_TURQUOISE: HSSFColor
        {
            public static short index = 0x29;
            public static short index2 = 0x1b;
            public static short[] triplet =
            {
                204, 255, 255
            };
            public static String hexString = "CCCC:FFFF:FFFF";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class PALE_BLUE
         *
         */

        public class PALE_BLUE: HSSFColor
        {
            public static short index = 0x2c;
            public static short[] triplet =
            {
                153, 204, 255
            };
            public static String hexString = "9999:CCCC:FFFF";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class LAVENDER
         *
         */

        public class LAVENDER: HSSFColor
        {
            public static short index = 0x2e;
            public static short[] triplet =
            {
                204, 153, 255
            };
            public static String hexString = "CCCC:9999:FFFF";


            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class WHITE
         *
         */

        public class WHITE: HSSFColor
        {
            public static short index = 0x9;
            public static short[] triplet =
            {
                255, 255, 255
            };
            public static String hexString = "FFFF:FFFF:FFFF";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class CORNFLOWER_BLUE
         */
        public class CORNFLOWER_BLUE: HSSFColor
        {
            public static short index = 0x18;
            public static short[] triplet =
            {
                153, 153, 255
            };
            public static String hexString = "9999:9999:FFFF";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }


        /**
         * Class LEMON_CHIFFON
         */
        public class LEMON_CHIFFON: HSSFColor
        {
            public static short index = 0x1a;
            public static short[] triplet =
            {
                255, 255, 204
            };

            public static String hexString = "FFFF:FFFF:CCCC";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class MAROON
         */
        public class MAROON: HSSFColor
        {
            public static short index = 0x19;
            public static short[] triplet =
            {
                127, 0, 0
            };
            public static String hexString = "8000:0:0";
            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class ORCHID
         */
        public class ORCHID: HSSFColor
        {
            public static short index = 0x1c;
            public static short[] triplet =
            {
                102, 0, 102
            };
            public static String hexString = "6666:0:6666";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class CORAL
         */
        public class CORAL : HSSFColor
        {
            public static short index = 0x1d;
            public static short[] triplet =
            {
                255, 128, 128
            };
            public static String hexString = "FFFF:8080:8080";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class ROYAL_BLUE
         */
        public class ROYAL_BLUE : HSSFColor
        {
            public static short index = 0x1e;
            public static short[] triplet =
        {
            0, 102, 204
        };
            public static String hexString = "0:6666:CCCC";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
            }
        }

        /**
         * Class LIGHT_CORNFLOWER_BLUE
         */
        public class LIGHT_CORNFLOWER_BLUE : HSSFColor
        {
            public static short index = 0x1f;
            public static short[] triplet =
        {
            204, 204, 255
        };
            public static String hexString = "CCCC:CCCC:FFFF";

            public override short GetIndex()
            {
                return index;
            }

            public override short[] GetTriplet()
            {
                return triplet;
            }

            public override String GetHexString()
            {
                return hexString; 
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
        public class AUTOMATIC : HSSFColor
        {
            private static HSSFColor instance = new AUTOMATIC();

            public static short index = 0x40;

            public override short[] GetTriplet()
            {
                return BLACK.triplet; 
            }

            public override String GetHexString()
            {
                return BLACK.hexString;
            }

            public static HSSFColor GetInstance()
            {
                return instance;
            }
        }
    }
}