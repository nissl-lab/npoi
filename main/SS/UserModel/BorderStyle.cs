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
        None = 0x0,

        /// <summary>
        /// Thin border
        /// </summary>
        Thin = 0x1,

        /// <summary>
        /// Medium border
        /// </summary>
        Medium = 0x2,

        /// <summary>
        /// dash border
        /// </summary>
        Dashed = 0x3,

        /// <summary>
        /// dot border
        /// </summary>
        Dotted = 0x4,

        /// <summary>
        /// Thick border
        /// </summary>
        Thick = 0x5,

        /// <summary>
        /// double-line border
        /// </summary>
        Double = 0x6,

        /// <summary>
        /// hair-line border
        /// </summary>
        Hair = 0x7,

        /// <summary>
        /// Medium dashed border
        /// </summary>
        MediumDashed = 0x8,

        /// <summary>
        /// dash-dot border
        /// </summary>
        DashDot = 0x9,

        /// <summary>
        /// medium dash-dot border
        /// </summary>
        MediumDashDot = 0xA,

        /// <summary>
        /// dash-dot-dot border
        /// </summary>
        DashDotDot = 0xB,

        /// <summary>
        /// medium dash-dot-dot border
        /// </summary>
        MediumDashDotDot = 0xC,

        /// <summary>
        /// slanted dash-dot border
        /// </summary>
        SlantedDashDot = 0xD
    }
}