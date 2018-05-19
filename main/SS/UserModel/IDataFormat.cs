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
    using System;

    public interface IDataFormat
    {
        /**
         * get the format index that matches the given format string.
         * Creates a new format if one is not found.  Aliases text to the proper format.
         * @param format string matching a built in format
         * @return index of format.
         */
        short GetFormat(String format);

        /**
         * get the format string that matches the given format index
         * @param index of a format
         * @return string represented at index of format or null if there is not a  format at that index
         */
        String GetFormat(short index);
    }
}