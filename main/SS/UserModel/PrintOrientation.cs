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
     * The enumeration value indicating the print orientation for a sheet.
     *
     * @author Gisella Bronzetti
     */
    public class PrintOrientation
    {

        /**
         * orientation not specified
         */
        public static PrintOrientation DEFAULT;
        /**
         * portrait orientation
         */
        public static PrintOrientation PORTRAIT;

        /**
         * landscape orientations
         */
        public static PrintOrientation LANDSCAPE;

        static PrintOrientation()
        { 
            _table = new PrintOrientation[4];
            DEFAULT = new PrintOrientation(1);
            PORTRAIT = new PrintOrientation(2);
            LANDSCAPE = new PrintOrientation(3);
        }

        private int orientation;

        private PrintOrientation(int orientation)
        {
            this.orientation = orientation;

            _table[this.Value] = this;
        }


        public int Value
        {
            get
            {
                return orientation;
            }
        }


        private static PrintOrientation[] _table;

        public static PrintOrientation ValueOf(int value)
        {
            return _table[value];
        }
    }
}
