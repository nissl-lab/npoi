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
    /// A property of a font that describes its general appearance.
    /// </summary>
    /// <remarks>
    /// @since POI 3.17-beta2
    /// </remarks>
    public enum FontFamily
    {
        /**
         * The default font is specified, which is implementation-dependent.
         */
        FF_DONTCARE =(0x00),
        /**
         * Fonts with variable stroke widths, which are proportional to the actual widths of
         * the glyphs, and which have serifs. "MS Serif" is an example.
         */
        FF_ROMAN= (0x01),
        /**
         * Fonts with variable stroke widths, which are proportional to the actual widths of the
         * glyphs, and which do not have serifs. "MS Sans Serif" is an example.
         */
        FF_SWISS= (0x02),
        /**
         * Fonts with constant stroke width, with or without serifs. Fixed-width fonts are
         * usually modern. "Pica", "Elite", and "Courier New" are examples.
         */
        FF_MODERN =(0x03),
        /**
         * Fonts designed to look like handwriting. "Script" and "Cursive" are examples.
         */
        FF_SCRIPT =(0x04),
        /**
         * Novelty fonts. "Old English" is an example.
         */
        FF_DECORATIVE =(0x05)
    }
    public static class FontFamilyExtension
    {
        private static Dictionary<int, FontFamily> values = new Dictionary<int, FontFamily>()
        {
        { 0, FontFamily.FF_DONTCARE },
        { 1, FontFamily.FF_ROMAN },
        { 2, FontFamily.FF_SWISS },
        { 3, FontFamily.FF_MODERN },
        { 4, FontFamily.FF_SCRIPT },
        { 5, FontFamily.FF_DECORATIVE },
    };
        public static int GetFlag(this FontFamily family)
        {
            return (int) family;
        }

        public static FontFamily? ValueOf(int nativeId)
        {
            if(values.TryGetValue(nativeId, out FontFamily family))
                return family;
            return null;
        }

        /**
         * Get FontFamily from combined native id
         *
         * @param pitchAndFamily The PitchFamily to decode.
         *
         * @return The resulting FontFamily
         */
        public static FontFamily? ValueOfPitchFamily(byte pitchAndFamily)
        {
            return ValueOf(pitchAndFamily >>> 4);
        }
    }
}
