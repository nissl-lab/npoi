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

namespace NPOI.HSSF.UserModel
{
    using System;
    using NPOI.SS.UserModel;
    using RecordExtendedColor = NPOI.HSSF.Record.Common.ExtendedColor;

    /**
     * The HSSF file format normally stores Color information in the
     *  Palette (see PaletteRecord), but for a few cases (eg Conditional
     *  Formatting, Sheet Extensions), this XSSF-style color record 
     *  can be used.
     */
    public class HSSFExtendedColor : ExtendedColor
    {
        private RecordExtendedColor color;

        public HSSFExtendedColor(NPOI.HSSF.Record.Common.ExtendedColor color)
        {
            this.color = color;
        }

        protected NPOI.HSSF.Record.Common.ExtendedColor GetExtendedColor()
        {
            return color;
        }

        public RecordExtendedColor ExtendedColor
        {
            get
            {
                return color;
            }
        }

        public override bool IsAuto
        {
            get
            {
                return color.type == RecordExtendedColor.TYPE_AUTO;
            }
            set
            {
                if (value)
                    color.type = RecordExtendedColor.TYPE_AUTO;
                else
                    color.type = RecordExtendedColor.TYPE_UNSET;
            }
        }
        public override bool IsIndexed
        {
            get
            {
                return color.type == RecordExtendedColor.TYPE_INDEXED;
            }
        }
        public override bool IsRGB
        {
            get
            {
                return color.type == RecordExtendedColor.TYPE_RGB;
            }
        }
        public override bool IsThemed
        {
            get
            {
                return color.type == RecordExtendedColor.TYPE_THEMED;
            }
        }

        public override short Index
        {
            get
            {
                return (short)color.ColorIndex;
            }
        }
        public override int Theme
        {
            get
            {
                return color.ThemeIndex;
            }
            set
            {
                color.ThemeIndex = value;
            }
        }

        public override byte[] RGB
        {
            get
            {
                // Trim trailing A
                byte[] rgb = new byte[3];
                byte[] rgba = color.RGBA;
                if (rgba == null) return null;
                Array.Copy(rgba, 0, rgb, 0, 3);
                return rgb;
            }
            set
            {
                if (value.Length == 3)
                {
                    byte[] rgba = new byte[4];
                    Array.Copy(value, 0, rgba, 0, 3);
                    unchecked
                    {
                        rgba[3] = (byte)-1;
                    }
                }
                else
                {
                    // Shuffle from ARGB to RGBA
                    byte a = value[0];
                    value[0] = value[1];
                    value[1] = value[2];
                    value[2] = value[3];
                    value[3] = a;
                    color.RGBA = (/*setter*/value);
                }
                color.Type = Record.Common.ExtendedColor.TYPE_RGB;
            }
        }
        public override byte[] ARGB
        {
            get
            {
                // Swap from RGBA to ARGB
                byte[] argb = new byte[4];
                byte[] rgba = color.RGBA;
                if (rgba == null) return null;
                Array.Copy(rgba, 0, argb, 1, 3);
                argb[0] = rgba[3];
                return argb;
            }
        }

        protected override byte[] StoredRGB
        {
            get
            {
                return ARGB;
            }
        }

        public void SetRGB(byte[] rgb)
        {
            
        }

        public override double Tint
        {
            get
            {
                return color.Tint;
            }
            set
            {
                color.Tint = (/*setter*/value);
            }
        }
    }

}