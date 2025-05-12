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

namespace NPOI.OOXML.XSSF.UserModel
{
    using NPOI.OpenXmlFormats.Spreadsheet;

    /// <summary>
    /// custom index color map, i.e. from the styles.xml definition
    /// </summary>
    public class CustomIndexedColorMap : IIndexedColorMap
    {

        private byte[][] colorIndex;

        /// <summary>
        /// </summary>
        /// <param name="colors">array of RGB triplets indexed by color index</param>
        private CustomIndexedColorMap(byte[][] colors)
        {
            this.colorIndex = colors;
        }

        public byte[] GetRGB(int index)
        {
            if(colorIndex == null || index < 0 || index >= colorIndex.Length)
                return null;
            return colorIndex[index];
        }

        /// <summary>
        /// <para>
        /// OOXML spec says if this exists it must have all indexes.
        /// </para>
        /// <para>
        /// From the OOXML Spec, Part 1, section 18.8.27:
        /// </para>
        /// <para>
        /// <i>
        /// This element contains a sequence of RGB color values that correspond to color indexes (zero-based). When
        /// using the default indexed color palette, the values are not written out, but instead are implied. When the color
        /// palette has been modified from default, then the entire color palette is written out.
        /// </i>
        /// </para>
        /// </summary>
        /// <param name="colors">CTColors from styles.xml possibly defining a custom color indexing scheme</param>
        /// <returns>custom indexed color map or null if none defined in the document</returns>
        public static CustomIndexedColorMap FromColors(CT_Colors colors)
        {
            if(colors == null || !colors.IsSetIndexedColors())
                return null;

            List<CT_RgbColor> rgbColorList = colors.indexedColors;
            byte[][] customColorIndex = new byte[rgbColorList.Count][];
            for(int i = 0; i < rgbColorList.Count; i++)
            {
                customColorIndex[i] = rgbColorList[i].rgb;
            }
            return new CustomIndexedColorMap(customColorIndex);
        }
    }
}


