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

namespace NPOI.HSSF.Record.Common
{
    using NPOI.Util;
    using System;
    using System.Text;


    /**
     * Title: CTColor (Extended Color) record part
     * <P>
     * The HSSF file format normally stores Color information in the
     *  Palette (see PaletteRecord), but for a few cases (eg Conditional
     *  Formatting, Sheet Extensions), this XSSF-style color record 
     *  can be used.
     */
    public class ExtendedColor: ICloneable {
        public static int TYPE_AUTO = 0;
        public static int TYPE_INDEXED = 1;
        public static int TYPE_RGB = 2;
        public static int TYPE_THEMED = 3;
        public static int TYPE_UNSET = 4;

        public static int THEME_DARK_1 = 0;
        public static int THEME_LIGHT_1 = 1;
        public static int THEME_DARK_2 = 2;
        public static int THEME_LIGHT_2 = 3;
        public static int THEME_ACCENT_1 = 4;
        public static int THEME_ACCENT_2 = 5;
        public static int THEME_ACCENT_3 = 6;
        public static int THEME_ACCENT_4 = 7;
        public static int THEME_ACCENT_5 = 8;
        public static int THEME_ACCENT_6 = 9;
        public static int THEME_HYPERLINK = 10;
        // This one is SheetEx only, not allowed in CFs
        public static int THEME_FOLLOWED_HYPERLINK = 11;

        public int type;

        // Type = Indexed
        public int colorIndex;
        // Type = RGB
        public byte[] rgba;
        // Type = Theme
        public int themeIndex;

        private double tint;

        public ExtendedColor() {
            this.type = TYPE_INDEXED;
            this.colorIndex = 0;
            this.tint = 0d;
        }
        public ExtendedColor(ILittleEndianInput in1) {
            type = in1.ReadInt();
            if (type == TYPE_INDEXED) {
                colorIndex = in1.ReadInt();
            } else if (type == TYPE_RGB) {
                rgba = new byte[4];
                in1.ReadFully(rgba);
            } else if (type == TYPE_THEMED) {
                themeIndex = in1.ReadInt();
            } else {
                // Ignored
                in1.ReadInt();
            }
            tint = in1.ReadDouble();
        }

        public int Type
        {
            get
            {
                return type;
            }
            set
            {
                this.type = value;
            }
        }

        /**
         * @return Palette color index, if type is {@link #TYPE_INDEXED}
         */
        public int ColorIndex
        {
            get
            {
                return colorIndex;
            }
            set
            {
                this.colorIndex = value;
            }
        }

        /**
         * @return Red Green Blue Alpha, if type is {@link #TYPE_RGB}
         */
        public byte[] RGBA
        {
            get { return rgba; }
            set { this.rgba = (value == null) ? null : (byte[])value.Clone(); }
        }

        /**
         * @return Theme color type index, eg {@link #THEME_DARK_1}, if type is {@link #TYPE_THEMED}
         */
        public int ThemeIndex
        {
            get { return themeIndex; }
            set { this.themeIndex = value; }
        }
        /**
         * @return Tint and Shade value, between -1 and +1
         */
        public double Tint
        {
            get { return tint; }
            set
            {
                if (tint < -1 || tint > 1)
                {
                    throw new ArgumentException("Tint/Shade must be between -1 and +1");
                }
                this.tint = value;
            }
        }

        public override String ToString() {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("    [Extended Color]\n");
            buffer.Append("          .type  = ").Append(type).Append("\n");
            buffer.Append("          .tint  = ").Append(tint).Append("\n");
            buffer.Append("          .c_idx = ").Append(colorIndex).Append("\n");
            buffer.Append("          .rgba  = ").Append(HexDump.ToHex(rgba)).Append("\n");
            buffer.Append("          .t_idx = ").Append(themeIndex).Append("\n");
            buffer.Append("    [/Extended Color]\n");
            return buffer.ToString();
        }

        public Object Clone() {
            ExtendedColor exc = new ExtendedColor();
            exc.type = type;
            exc.tint = tint;
            if (type == TYPE_INDEXED) {
                exc.colorIndex = colorIndex;
            } else if (type == TYPE_RGB) {
                exc.rgba = new byte[4];
                Array.Copy(rgba, 0, exc.rgba, 0, 4);
            } else if (type == TYPE_THEMED) {
                exc.themeIndex = themeIndex;
            }
            return exc;
        }

        public int DataLength
        {
            get
            {
                return 4 + 4 + 8;
            }
        }

        public void Serialize(ILittleEndianOutput out1) {
            out1.WriteInt(type);
            if (type == TYPE_INDEXED) {
                out1.WriteInt(colorIndex);
            } else if (type == TYPE_RGB) {
                out1.Write(rgba);
            } else if (type == TYPE_THEMED) {
                out1.WriteInt(themeIndex);
            } else {
                out1.WriteInt(0);
            }
            out1.WriteDouble(tint);
        }
    }
}