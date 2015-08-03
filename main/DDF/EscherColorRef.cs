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

namespace NPOI.DDF
{
    using System;
    using NPOI.Util;
    using System.Diagnostics;
    public class SysIndexSource
    {
        internal static SysIndexSource[] Values()
        {
            return new SysIndexSource[] {
                SysIndexSource.FILL_COLOR,
                SysIndexSource.LINE_OR_FILL_COLOR,
                SysIndexSource.SHADOW_COLOR,
                SysIndexSource.CURRENT_OR_LAST_COLOR,
                SysIndexSource.FILL_BACKGROUND_COLOR,
                SysIndexSource.LINE_BACKGROUND_COLOR,
                SysIndexSource.FILL_OR_LINE_COLOR
            };
        }
        /* Use the fill color of the shape. */
        public static SysIndexSource FILL_COLOR = new SysIndexSource(0xF0),
        /* If the shape Contains a line, use the line color of the shape. Otherwise, use the fill color. */
        LINE_OR_FILL_COLOR = new SysIndexSource(0xF1),
        /* Use the line color of the shape. */
        LINE_COLOR = new SysIndexSource(0xF2),
            /* Use the shadow color of the shape. */
        SHADOW_COLOR = new SysIndexSource(0xF3),
            /* Use the current, or last-used, color. */
        CURRENT_OR_LAST_COLOR = new SysIndexSource(0xF4),
            /* Use the fill background color of the shape. */
        FILL_BACKGROUND_COLOR = new SysIndexSource(0xF5),
            /* Use the line background color of the shape. */
        LINE_BACKGROUND_COLOR = new SysIndexSource(0xF6),
            /* If the shape Contains a Fill, use the fill color of the shape. Otherwise, use the line color. */
        FILL_OR_LINE_COLOR = new SysIndexSource(0xF7)
        ;
        internal int value;
        internal SysIndexSource(int value) { this.value = value; }
    }
    /**
         * The following enum specifies values that indicate special procedural properties that
         * are used to modify the color components of another color. These values are combined with
         * those of the {@link SysIndexSource} enum or with a user-specified color.
         * The first six values are mutually exclusive.
         */
    public class SysIndexProcedure
    {
        internal static SysIndexProcedure[] Values()
        {
            return new SysIndexProcedure[] { 
                SysIndexProcedure.DARKEN_COLOR,
                SysIndexProcedure.LIGHTEN_COLOR,
                SysIndexProcedure.ADD_GRAY_LEVEL,
                SysIndexProcedure.SUB_GRAY_LEVEL,
                SysIndexProcedure.REVERSE_GRAY_LEVEL,
                SysIndexProcedure.THRESHOLD,
                SysIndexProcedure.INVERT_AFTER,
                SysIndexProcedure.INVERT_HIGHBIT_AFTER
            };
        }
        /*
         * Darken the color by the value that is specified in the blue field.
         * A blue value of 0xFF specifies that the color is to be left unChanged,
         * whereas a blue value of 0x00 specifies that the color is to be completely darkened.
         */
        public static SysIndexProcedure DARKEN_COLOR = new SysIndexProcedure(0x01),
            /*
             * Lighten the color by the value that is specified in the blue field.
             * A blue value of 0xFF specifies that the color is to be left unChanged,
             * whereas a blue value of 0x00 specifies that the color is to be completely lightened.
             */
        LIGHTEN_COLOR = new SysIndexProcedure(0x02),
            /*
             * Add a gray level RGB value. The blue field Contains the gray level to Add:
             * NewColor = SourceColor + gray
             */
        ADD_GRAY_LEVEL = new SysIndexProcedure(0x03),
            /*
             * Subtract a gray level RGB value. The blue field Contains the gray level to subtract:
             * NewColor = SourceColor - gray
             */
        SUB_GRAY_LEVEL = new SysIndexProcedure(0x04),
            /*
             * Reverse-subtract a gray level RGB value. The blue field Contains the gray level from
             * which to subtract:
             * NewColor = gray - SourceColor
             */
        REVERSE_GRAY_LEVEL = new SysIndexProcedure(0x05),
            /*
             * If the color component being modified is less than the parameter Contained in the blue
             * field, Set it to the minimum intensity. If the color component being modified is greater
             * than or equal to the parameter, Set it to the maximum intensity.
             */
        THRESHOLD = new SysIndexProcedure(0x06),
            /*
             * After making other modifications, invert the color.
             * This enum value is only for documentation and won't be directly returned.
             */
        INVERT_AFTER = new SysIndexProcedure(0x20),
            /*
             * After making other modifications, invert the color by toggling just the high bit of each
             * color channel.
             * This enum value is only for documentation and won't be directly returned.
             */
        INVERT_HIGHBIT_AFTER = new SysIndexProcedure(0x40)
        ;
        internal BitField mask;
        internal SysIndexProcedure(int mask)
        {
            this.mask = new BitField(mask);
        }
    }
    /**
     * An OfficeArtCOLORREF structure entry which also handles color extension opid data
     */
    public class EscherColorRef
    {
        int opid = -1;
        int colorRef = 0;

        /*
         * A bit that specifies whether the system color scheme will be used to determine the color. 
         * A value of 0x1 specifies that green and red will be treated as an unsigned 16-bit index
         * into the system color table. Values less than 0x00F0 map directly to system colors.
         */
        private static BitField FLAG_SYS_INDEX = new BitField(0x10000000);

        /*
         * A bit that specifies whether the current application-defined color scheme will be used
         * to determine the color. A value of 0x1 specifies that red will be treated as an index
         * into the current color scheme table. If this value is 0x1, green and blue MUST be 0x00.
         */
        private static BitField FLAG_SCHEME_INDEX = new BitField(0x08000000);

        /*
         * A bit that specifies whether the color is a standard RGB color.
         * 0x0 : The RGB color MAY use halftone dithering to display.
         * 0x1 : The color MUST be a solid color.
         */
        private static BitField FLAG_SYSTEM_RGB = new BitField(0x04000000);

        /*
         * A bit that specifies whether the current palette will be used to determine the color.
         * A value of 0x1 specifies that red, green, and blue contain an RGB value that will be
         * matched in the current color palette. This color MUST be solid.
         */
        private static BitField FLAG_PALETTE_RGB = new BitField(0x02000000);

        /*
         * A bit that specifies whether the current palette will be used to determine the color.
         * A value of 0x1 specifies that green and red will be treated as an unsigned 16-bit index into 
         * the current color palette. This color MAY be dithered. If this value is 0x1, blue MUST be 0x00.
         */
        private static BitField FLAG_PALETTE_INDEX = new BitField(0x01000000);

        /*
         * An unsigned integer that specifies the intensity of the blue color channel. A value
         * of 0x00 has the minimum blue intensity. A value of 0xFF has the maximum blue intensity.
         */
        private static BitField FLAG_BLUE = new BitField(0x00FF0000);

        /*
         * An unsigned integer that specifies the intensity of the green color channel. A value
         * of 0x00 has the minimum green intensity. A value of 0xFF has the maximum green intensity.
         */
        private static BitField FLAG_GREEN = new BitField(0x0000FF00);

        /*
         * An unsigned integer that specifies the intensity of the red color channel. A value
         * of 0x00 has the minimum red intensity. A value of 0xFF has the maximum red intensity.
         */
        private static BitField FLAG_RED = new BitField(0x000000FF);

        public EscherColorRef(int colorRef)
        {
            this.colorRef = colorRef;
        }

        public EscherColorRef(byte[] source, int start, int len)
        {
            Debug.Assert(len == 4 || len == 6);

            int offset = start;
            if (len == 6)
            {
                opid = LittleEndian.GetUShort(source, offset);
                offset += 2;
            }
            colorRef = LittleEndian.GetInt(source, offset);
        }

        public bool HasSysIndexFlag()
        {
            return FLAG_SYS_INDEX.IsSet(colorRef);
        }

        public void SetSysIndexFlag(bool flag)
        {
            FLAG_SYS_INDEX.SetBoolean(colorRef, flag);
        }

        public bool HasSchemeIndexFlag()
        {
            return FLAG_SCHEME_INDEX.IsSet(colorRef);
        }

        public void SetSchemeIndexFlag(bool flag)
        {
            FLAG_SCHEME_INDEX.SetBoolean(colorRef, flag);
        }

        public bool HasSystemRGBFlag()
        {
            return FLAG_SYSTEM_RGB.IsSet(colorRef);
        }

        public void SetSystemRGBFlag(bool flag)
        {
            FLAG_SYSTEM_RGB.SetBoolean(colorRef, flag);
        }

        public bool HasPaletteRGBFlag()
        {
            return FLAG_PALETTE_RGB.IsSet(colorRef);
        }

        public void SetPaletteRGBFlag(bool flag)
        {
            FLAG_PALETTE_RGB.SetBoolean(colorRef, flag);
        }

        public bool HasPaletteIndexFlag()
        {
            return FLAG_PALETTE_INDEX.IsSet(colorRef);
        }

        public void SetPaletteIndexFlag(bool flag)
        {
            FLAG_PALETTE_INDEX.SetBoolean(colorRef, flag);
        }

        public int[] GetRGB()
        {
            int[] rgb = {
            FLAG_RED.GetValue(colorRef),
            FLAG_GREEN.GetValue(colorRef),
            FLAG_BLUE.GetValue(colorRef)
        };
            return rgb;
        }

        /**
         * @return {@link SysIndexSource} if {@link #hasSysIndexFlag()} is {@code true}, otherwise null
         */
        public SysIndexSource GetSysIndexSource()
        {
            if (!HasSysIndexFlag()) return null;
            int val = FLAG_RED.GetValue(colorRef);
            foreach (SysIndexSource sis in SysIndexSource.Values())
            {
                if (sis.value == val) return sis;
            }
            return null;
        }

        /**
         * Return the {@link SysIndexProcedure} - for invert flag use {@link #getSysIndexInvert()}
         * @return {@link SysIndexProcedure} if {@link #hasSysIndexFlag()} is {@code true}, otherwise null
         */
        public SysIndexProcedure GetSysIndexProcedure()
        {
            if (!HasSysIndexFlag()) return null;
            int val = FLAG_RED.GetValue(colorRef);
            foreach (SysIndexProcedure sip in SysIndexProcedure.Values())
            {
                if (sip == SysIndexProcedure.INVERT_AFTER || sip == SysIndexProcedure.INVERT_HIGHBIT_AFTER) continue;
                if (sip.mask.IsSet(val)) return sip;
            }
            return null;
        }

        /**
         * @return 0 for no invert flag, 1 for {@link SysIndexProcedure#INVERT_AFTER} and
         * 2 for {@link SysIndexProcedure#INVERT_HIGHBIT_AFTER} 
         */
        public int GetSysIndexInvert()
        {
            if (!HasSysIndexFlag()) return 0;
            int val = FLAG_GREEN.GetValue(colorRef);
            if ((SysIndexProcedure.INVERT_AFTER.mask.IsSet(val))) return 1;
            if ((SysIndexProcedure.INVERT_HIGHBIT_AFTER.mask.IsSet(val))) return 2;
            return 0;
        }

        /**
         * @return index of the scheme color or -1 if {@link #hasSchemeIndexFlag()} is {@code false}
         * 
         * @see NPOI.HSLF.Record.ColorSchemeAtom#getColor(int)
         */
        public int GetSchemeIndex()
        {
            if (!HasSchemeIndexFlag()) return -1;
            return FLAG_RED.GetValue(colorRef);
        }

        /**
         * @return index of current palette (color) or -1 if {@link #hasPaletteIndexFlag()} is {@code false}
         */
        public int GetPaletteIndex()
        {
            if (!HasPaletteIndexFlag()) return -1;
            return (FLAG_GREEN.GetValue(colorRef) << 8) & FLAG_RED.GetValue(colorRef);
        }
    }
    
}