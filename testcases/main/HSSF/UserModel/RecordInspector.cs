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


namespace TestCases.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.UserModel;
    using System.Collections;

    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;

    /**
     * Test utility class to get {@link Record}s out HSSF objects
     * 
     * @author Josh Micich
     */
    public class RecordInspector
    {

        private RecordInspector()
        {
            // no instances of this class
        }

        public class RecordCollector : RecordVisitor
        {

            private ArrayList _list;

            public RecordCollector()
            {
                _list = new ArrayList(128);
            }

            public void VisitRecord(Record r)
            {
                _list.Add(r);
            }

            public Record[] Records
            {
                get
                {
                    return (Record[])_list.ToArray(typeof(Record));
                }
            }
        }

        /**
         * @param streamOffset start position for serialization. This affects values in some
         *         records such as INDEX, but most callers will be OK to pass zero.
         * @return the {@link Record}s (in order) which will be output when the
         *         specified sheet is serialized
         */
        public static Record[] GetRecords(NPOI.SS.UserModel.ISheet hSheet, int streamOffset)
        {
            RecordCollector rc = new RecordCollector();
            ((HSSFSheet)hSheet).Sheet.VisitContainedRecords(rc, streamOffset);
            return rc.Records;
        }
    }
}