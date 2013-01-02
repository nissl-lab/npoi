
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

namespace NPOI.HSSF.EventUserModel
{
    using NPOI.HSSF.Record;

    /// <summary>
    /// Interface for use with the HSSFRequest and HSSFEventFactory.  Users should Create
    /// a listener supporting this interface and register it with the HSSFRequest (associating
    /// it with Record SID's).
    /// @author  acoliver@apache.org
    /// </summary>
    public interface IHSSFListener
    {

        /// <summary>
        /// Process an HSSF Record. Called when a record occurs in an HSSF file.
        /// </summary>
        /// <param name="record">The record.</param>
        void ProcessRecord(Record record);
    }
}