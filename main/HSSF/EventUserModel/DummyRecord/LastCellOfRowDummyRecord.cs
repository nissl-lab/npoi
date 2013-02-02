/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.EventUserModel.DummyRecord
{
    /**
     * A dummy record to indicate that we've now had the last
     *  cell record for this row.
     */
    public class LastCellOfRowDummyRecord : DummyRecordBase
    {
        private int row;
        private int lastColumnNumber;

        public LastCellOfRowDummyRecord(int row, int lastColumnNumber)
        {
            this.row = row;
            this.lastColumnNumber = lastColumnNumber;
        }

        /**
         * Returns the (0 based) number of the row we are
         *  currently working on.
         */
        public int Row
        {
            get { return row; }
        }
        /**
         * Returns the (0 based) number of the last column
         *  seen for this row. You should have alReady been
         *  called with that record.
         * This Is -1 in the case of there being no columns
         *  for the row.
         */
        public int LastColumnNumber
        {
            get { return lastColumnNumber; }
        }
    }
}