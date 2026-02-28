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
     * Defines the font scheme to which this font belongs.
     * When a font defInition is part of a theme defInition, then the font is categorized as either a major or minor font scheme component.
     * When a new theme is chosen, every font that is part of a theme defInition is updated to use the new major or minor font defInition for that
     * theme.
     * Usually major fonts are used for styles like headings, and minor fonts are used for body and paragraph text.
     *
     * @author Gisella Bronzetti
     */
    public class FontScheme
    {
        public static readonly FontScheme NONE = new FontScheme(1);
        public static readonly FontScheme MAJOR = new FontScheme(2);
        public static readonly FontScheme MINOR = new FontScheme(3);

        private int value;

        private FontScheme(int val)
        {
            value = val;
        }

        public int Value
        {
            get
            {
                return value;
            }
        }

        public static FontScheme ValueOf(int value)
        {
            switch (value)
            {
                case 1: return NONE;
                case 2: return MAJOR;
                case 3: return MINOR;
            }
            return NONE;
        }
    }

}