
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */


namespace NPOI.HSSF.Record
{
    using System;

    /// <summary>
    /// Used by records to indicate invalid format/data.
    /// </summary>
    public class RecordFormatException : Exception
    {
        public RecordFormatException(String exception):base(exception)
        {
        }

        public RecordFormatException(String exception, Exception thr):base(exception, thr)
        {
            
        }

        public RecordFormatException(Exception thr):base("",thr)
        {
            
        }
    }
}