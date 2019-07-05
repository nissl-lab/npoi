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


namespace NPOI.HWPF.Model
{
    using System;
    using System.Collections;
    using NPOI.Util;
    using NPOI.HWPF.Model.IO;
    /**
     * String table containing the history of the last few revisions ("saves") of the document.
     * Read-only for the time being.
     * 
     * @author Daniel Noll
     */
    public class SavedByTable
    {
        /**
         * A value that I don't know what it does, but Is maintained for accuracy.
         */
        private short unknownValue = -1;

        /**
         * Array of entries.
         */
        private SavedByEntry[] entries;

        /**
         * Constructor to read the table from the table stream.
         *
         * @param tableStream the table stream.
         * @param offset the offset into the byte array.
         * @param size the size of the table in the byte array.
         */
        public SavedByTable(byte[] tableStream, int offset, int size)
        {
            // Read the value that I don't know what it does. :-)
            unknownValue = LittleEndian.GetShort(tableStream, offset);
            offset += 2;

            // The stored int Is the number of strings, and there are two strings per entry.
            int numEntries = LittleEndian.GetInt(tableStream, offset) / 2;
            offset += 4;

            entries = new SavedByEntry[numEntries];
            for (int i = 0; i < numEntries; i++)
            {
                int len = LittleEndian.GetShort(tableStream, offset);
                offset += 2;
                String userName = StringUtil.GetFromUnicodeLE(tableStream, offset, len);
                offset += len * 2;
                len = LittleEndian.GetShort(tableStream, offset);
                offset += 2;
                String saveLocation = StringUtil.GetFromUnicodeLE(tableStream, offset, len);
                offset += len * 2;

                entries[i] = new SavedByEntry(userName, saveLocation);
            }
        }

        /**
         * Gets the entries.  The returned list cannot be modified.
         *
         * @return the list of entries.
         */
        public IList GetEntries()
        {
            return Arrays.AsList(entries);
        }

        /**
         * Writes this table to the table stream.
         *
         * @param tableStream the table stream to write to.
         * @throws IOException if an error occurs while writing.
         */
        public void WriteTo(HWPFStream tableStream)
        {
            byte[] header = new byte[6];
            LittleEndian.PutShort(header, 0, unknownValue);
            LittleEndian.PutInt(header, 2, entries.Length * 2);
            tableStream.Write(header);

            for (int i = 0; i < entries.Length; i++)
            {
                WriteStringValue(tableStream, entries[i].GetUserName());
                WriteStringValue(tableStream, entries[i].GetSaveLocation());
            }
        }

        private void WriteStringValue(HWPFStream tableStream, String value)
        {
            byte[] buf = new byte[value.Length * 2 + 2];
            LittleEndian.PutShort(buf, 0, (short)value.Length);
            StringUtil.PutUnicodeLE(value, buf, 2);
            tableStream.Write(buf);
        }
    }
}
