/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.Common.UserModel.Fonts
{
    /// <summary>
    /// A property of a font that describes the pitch, of the characters.
    /// </summary>
    /// <remarks>
    /// @since POI 3.17-beta2
    /// </remarks>
    public enum FontPitch
    {
        /**
         * The default pitch, which is implementation-dependent.
         */
        DEFAULT = (0x00),
        /**
         * A fixed pitch, which means that all the characters in the font occupy the same
         * width when output in a string.
         */
        FIXED = (0x01),
        /**
         * A variable pitch, which means that the characters in the font occupy widths
         * that are proportional to the actual widths of the glyphs when output in a string. For example,
         * the "i" and space characters usually have much smaller widths than a "W" or "O" character.
         */
        VARIABLE = (0x02)
    }
    public static class FontPitchExtension
    {

        private static Dictionary<int, FontPitch> values = new Dictionary<int, FontPitch>()
            {
            { 0, FontPitch.DEFAULT }, { 1, FontPitch.FIXED }, { 2, FontPitch.VARIABLE }
        };

        public static int GetNativeId(this FontPitch value)
        {
            return (int) value;
        }

        public static FontPitch? ValueOf(int flag)
        {
            if(values.TryGetValue(flag, out FontPitch value))
                return value;
            return null;
        }

        /**
         * Combine pitch and family to native id
         * 
         * @see <a href="https://msdn.microsoft.com/en-us/library/dd145037.aspx">LOGFONT structure</a>
         *
         * @param pitch The pitch-value, cannot be null
         * @param family The family-value, cannot be null
         *
         * @return The resulting combined byte-value with pitch and family encoded into one byte
         */
        public static byte GetNativeId(FontPitch pitch, FontFamily family)
        {
            return (byte) (pitch.GetNativeId() | (family.GetFlag() << 4));
        }

        /**
         * Get FontPitch from native id
         *
         * @param pitchAndFamily The combined byte value for pitch and family
         *
         * @return The resulting FontPitch enumeration value
         */
        public static FontPitch? ValueOfPitchFamily(byte pitchAndFamily)
        {
            return ValueOf(pitchAndFamily & 0x3);
        }
    }
}
