
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

namespace NPOI.DDF
{

    /// <summary>
    /// The escher record factory interface allows for the creation of escher
    /// records from a pointer into a data array.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public interface IEscherRecordFactory
    {
        /// <summary>
        /// Create a new escher record from the data provided.  Does not attempt
        /// to Fill the contents of the record however.
        /// </summary>
        /// <param name="data">The data.</param>
        /// <param name="offset">The off set.</param>
        /// <returns></returns>
        EscherRecord CreateRecord(byte[] data, int offset);
    }
}