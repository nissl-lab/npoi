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
     * Specifies printed page order.
     *
     * @author Gisella Bronzetti
     */
    public class PageOrder
    {

        /**
         * Order pages vertically first, then move horizontally.
         */
        public static PageOrder DOWN_THEN_OVER;
        /**
         * Order pages horizontally first, then move vertically
         */
        public static PageOrder OVER_THEN_DOWN;


        private int order;

        static PageOrder()
        { 
            _table = new PageOrder[3];
            DOWN_THEN_OVER = new PageOrder(1);
            OVER_THEN_DOWN = new PageOrder(2);
        }

        private PageOrder(int order)
        {
            this.order = order;
            _table[order] = this;
        }

        public int Value
        {
            get
            {
                return order;
            }
        }

        private static PageOrder[] _table;

        public static PageOrder ValueOf(int value)
        {
            return _table[value];
        }
    }

}
