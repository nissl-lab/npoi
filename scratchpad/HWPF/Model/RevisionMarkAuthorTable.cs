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


namespace NPOI.HWPF.Model
{
    using NPOI.Util;
    using System;
    using System.Collections.Generic;
    using NPOI.HWPF.Model.IO;

    /**
     * String table Containing the names of authors of revision marks, e-mails and
     * comments in this document.
     * 
     * @author Ryan Lauck
     */
    public class RevisionMarkAuthorTable
    {
        /**
         * must be 0xFFFF
         */
        private short fExtend = unchecked((short)0xFFFF);

        /**
         * the number of entries in the table
         */
        private short cData = 0;

        /**
         * must be 0
         */
        private short cbExtra = 0;

        /**
         * Array of entries.
         */
        private String[] entries;

        /**
         * Constructor to read the table from the table stream.
         * 
         * @param tableStream the table stream.
         * @param offset the offset into the byte array.
         * @param size the size of the table in the byte array.
         */
        public RevisionMarkAuthorTable(byte[] tableStream, int offset, int size)
        {
            // Read fExtend - it isn't used
            fExtend = LittleEndian.GetShort(tableStream, offset);
            if (fExtend != unchecked((short)0xFFFF))
            {
                //TODO: throw an exception here?
            }
            offset += 2;

            // Read the number of entries
            cData = LittleEndian.GetShort(tableStream, offset);
            offset += 2;

            // Read cbExtra - it isn't used
            cbExtra = LittleEndian.GetShort(tableStream, offset);
            if (cbExtra != 0)
            {
                //TODO: throw an exception here?
            }
            offset += 2;

            entries = new String[cData];
            for (int i = 0; i < cData; i++)
            {
                int len = LittleEndian.GetShort(tableStream, offset);
                offset += 2;

                String name = StringUtil.GetFromUnicodeLE(tableStream, offset, len);
                offset += len * 2;

                entries[i] = name;
            }
        }

        /**
         * Gets the entries. The returned list cannot be modified.
         * 
         * @return the list of entries.
         */
        public List<String> GetEntries()
        {
            return new List<String>(entries);
        }

        /**
         * Get an author by its index.  Returns null if it does not exist.
         * 
         * @return the revision mark author
         */
        public String GetAuthor(int index)
        {
            String auth = null;
            if (index >= 0 && index < entries.Length)
            {
                auth = entries[index];
            }
            return auth;
        }

        /**
         * Gets the number of entries.
         * 
         * @return the number of entries.
         */
        public int GetSize()
        {
            return cData;
        }

        /**
         * Writes this table to the table stream.
         * 
         * @param tableStream  the table stream to write to.
         * @throws IOException  if an error occurs while writing.
         */
        public void WriteTo(HWPFStream tableStream)
        {
            byte[] header = new byte[6];
            LittleEndian.PutShort(header, 0, fExtend);
            LittleEndian.PutShort(header, 2, cData);
            LittleEndian.PutShort(header, 4, cbExtra);
            tableStream.Write(header);

            foreach (String name in entries)
            {
                byte[] buf = new byte[name.Length * 2 + 2];
                LittleEndian.PutShort(buf, 0, (short)name.Length);
                StringUtil.PutUnicodeLE(name, buf, 2);
                tableStream.Write(buf);
            }
        }

    }
}

