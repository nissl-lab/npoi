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

namespace NPOI.HSSF.Record
{

    using System;
    using System.Text;

    [Obsolete("Not found in poi,is it useful?")]
    public interface CustomField
           : ICloneable
    {
        /**
         * @return  The size of this field in bytes.  This operation Is not valid
         *          Until after the call to <c>FillField()</c>
         */
        int Size { get; }

        /**
         * Populates this fields data from the byte array passed in1.
         * @param in the RecordInputstream to Read the record from
         */
        int FillField(RecordInputStream in1);

        /**
         * Appends the string representation of this field to the supplied
         * StringBuilder.
         *
         * @param str   The string buffer to Append to.
         */
        void ToString(StringBuilder str);

        /**
         * Converts this field to it's byte array form.
         * @param offset    The offset into the byte array to start writing to.
         * @param data      The data array to Write to.
         * @return  The number of bytes written.
         */
        int SerializeField(int offset, byte[] data);


    }
}