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

    /**
     * The enumeration value indicating the line style of a border in a cell,
     * i.e., whether it is borded dash dot, dash dot dot, dashed, dotted, double, hair, medium, 
     * medium dash dot, medium dash dot dot, medium dashed, none, slant dash dot, thick or thin.
     */
    public enum BorderStyle
    {

        /**
         * No border
         */

        NONE,

        /**
         * Thin border
         */

        THIN,

        /**
         * Medium border
         */

        MEDIUM,

        /**
         * dash border
         */

        DASHED,

        /**
         * dot border
         */

        HAIR,

        /**
         * Thick border
         */

        THICK,

        /**
         * double-line border
         */

        DOUBLE,

        /**
         * hair-line border
         */

        DOTTED,

        /**
         * Medium dashed border
         */

        MEDIUM_DASHED,

        /**
         * dash-dot border
         */

        DASH_DOT,

        /**
         * medium dash-dot border
         */

        MEDIUM_DASH_DOT,

        /**
         * dash-dot-dot border
         */

        DASH_DOT_DOT,

        /**
         * medium dash-dot-dot border
         */

        MEDIUM_DASH_DOT_DOTC,

        /**
         * slanted dash-dot border
         */

        SLANTED_DASH_DOT

    }

}