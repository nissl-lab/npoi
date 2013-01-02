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
     * These enumerations specify how cell comments shall be displayed for paper printing purposes.
     *
     * @author Gisella Bronzetti
     */
    public class PrintCellComments
    {

        /**
         * Do not print cell comments.
         */
        public static PrintCellComments NONE;
        /**
         * Print cell comments as displayed.
         */
        public static PrintCellComments AS_DISPLAYED;
        /**
         * Print cell comments at end of document.
         */
        public static PrintCellComments AT_END;


        static PrintCellComments()
        { 
            _table= new PrintCellComments[4];
            NONE = new PrintCellComments(1);
            AS_DISPLAYED = new PrintCellComments(2);
            AT_END = new PrintCellComments(3);
        }

        private int comments;

        private PrintCellComments(int comments)
        {
            this.comments = comments;
            _table[this.Value] = this;
        }

        public int Value
        {
            get
            {
                return comments;
            }
        }

        private static PrintCellComments[] _table;

        public static PrintCellComments ValueOf(int value)
        {
            return _table[value];
        }
    }

}