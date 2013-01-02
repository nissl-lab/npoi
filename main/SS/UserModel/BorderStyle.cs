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

    /// <summary>
    /// The enumeration value indicating the line style of a border in a cell
    /// </summary>
    public enum BorderStyle : short
    {
        /// <summary>
        /// No border
        /// </summary>
        NONE = 0x0,

        /// <summary>
        /// Thin border
        /// </summary>
        THIN = 0x1,

        /// <summary>
        /// Medium border
        /// </summary>
        MEDIUM = 0x2,

        /// <summary>
        /// dash border
        /// </summary>
        DASHED = 0x3,

        /// <summary>
        /// dot border
        /// </summary>
        DOTTED = 0x4,

        /// <summary>
        /// Thick border
        /// </summary>
        THICK = 0x5,

        /// <summary>
        /// double-line border
        /// </summary>
        DOUBLE = 0x6,

        /// <summary>
        /// hair-line border
        /// </summary>
        HAIR = 0x7,

        /// <summary>
        /// Medium dashed border
        /// </summary>
        MEDIUM_DASHED = 0x8,

        /// <summary>
        /// dash-dot border
        /// </summary>
        DASH_DOT = 0x9,

        /// <summary>
        /// medium dash-dot border
        /// </summary>
        MEDIUM_DASH_DOT = 0xA,

        /// <summary>
        /// dash-dot-dot border
        /// </summary>
        DASH_DOT_DOT = 0xB,

        /// <summary>
        /// medium dash-dot-dot border
        /// </summary>
        MEDIUM_DASH_DOT_DOT = 0xC,

        /// <summary>
        /// slanted dash-dot border
        /// </summary>
        SLANTED_DASH_DOT = 0xD

    }

}