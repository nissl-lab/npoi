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

namespace NPOI.HSSF.EventModel
{
    using NPOI.HSSF.Record;

    /**
     * An ERFListener Is registered with the EventRecordFactory.
     * An ERFListener listens for Records coming from the stream
     * via the EventRecordFactory
     * 
     * @see EventRecordFactory
     * @author Andrew C. Oliver acoliver@apache.org
     */
    public interface IERFListener
    {
        /**
         * Process a Record.  This method Is called by the 
         * EventRecordFactory when a record Is returned.
         * @return bool specifying whether the effort was a success.
         */
        bool ProcessRecord(Record rec);
    }
}