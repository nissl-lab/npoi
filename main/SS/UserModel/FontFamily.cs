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
     * The font family this font belongs to. A font family is a set of fonts having common stroke width and serif
     * characteristics. The font name overrides when there are conflicting values.
     *
     * @author Gisella Bronzetti
     */
    public class FontFamily
    {
        public static readonly FontFamily NOT_APPLICABLE = new FontFamily(0);
        public static readonly FontFamily ROMAN = new FontFamily(1);
        public static readonly FontFamily SWISS = new FontFamily(2);
        public static readonly FontFamily MODERN = new FontFamily(3);
        public static readonly FontFamily SCRIPT = new FontFamily(4);
        public static readonly FontFamily DECORATIVE = new FontFamily(5);

        private int family;

        private FontFamily(int value)
        {
            family = value;
        }

        /**
         * Returns index of this font family
         *
         * @return index of this font family
         */
        public int Value
        {
            get
            {
                return family;
            }
        }

        public static FontFamily ValueOf(int family)
        {
            switch(family)
            {
                case 0: return NOT_APPLICABLE;
                case 1: return ROMAN;
                case 2: return SWISS;
                case 3: return MODERN;
                case 4: return SCRIPT;
                case 5: return DECORATIVE;
            }
            return NOT_APPLICABLE;
        }
    }
}

