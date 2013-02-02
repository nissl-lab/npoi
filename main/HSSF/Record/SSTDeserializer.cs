
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
    using NPOI.Util;

    /**
     * Handles the task of deserializing a SST string.  The two main entry points are
     *
     * @author Glen Stampoultzis (glens at apache.org)
     * @author Jason Height (jheight at apache.org)
     */
    public class SSTDeserializer
    {

        private IntMapper<UnicodeString> strings;

        public SSTDeserializer(IntMapper<UnicodeString> strings)
        {
            this.strings = strings;
        }

        /**
         * This Is the starting point where strings are constructed.  Note that
         * strings may span across multiple continuations. Read the SST record
         * carefully before beginning to hack.
         */
        public void ManufactureStrings(int stringCount, RecordInputStream in1)
        {
            for (int i = 0; i < stringCount; i++)
            {
                // Extract exactly the count of strings from the SST record.
                UnicodeString str;
                if (in1.Available() == 0 && !in1.HasNextRecord)
                {
                    System.Console.WriteLine("Ran out of data before creating all the strings! String at index " + i + "");
                    str = new UnicodeString("");
                }
                else
                {
                    str = new UnicodeString(in1);
                }
                AddToStringTable(strings, str);
            }
        }

        static public void AddToStringTable(IntMapper<UnicodeString> strings, UnicodeString str)
        {
            strings.Add(str);
        }
    }
}