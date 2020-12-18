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
using NPOI.Util;
using System;
using NPOI.HWPF.Model.IO;
namespace NPOI.HWPF.Model
{
    /**
     * Utils for storing and Reading "STring TaBle stored in File"
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     */

    public class SttbfUtils
    {
        public static String[] Read(byte[] data, int startOffset)
        {
            short ffff = LittleEndian.GetShort(data, startOffset);

            if (ffff != unchecked((short)0xffff))
            {
                // Non-extended character Pascal strings
                throw new NotSupportedException(
                        "Non-extended character Pascal strings are not supported right now. "
                                + "Please, contact POI developers for update.");
            }

            // strings are extended character strings
            int offset = startOffset + 2;
            int numEntries = LittleEndian.GetInt(data, offset);
            offset += 4;

            String[] entries = new String[numEntries];
            for (int i = 0; i < numEntries; i++)
            {
                int len = LittleEndian.GetShort(data, offset);
                offset += 2;
                String value = StringUtil.GetFromUnicodeLE(data, offset, len);
                offset += len * 2;
                entries[i] = value;
            }
            return entries;
        }

        public static int Write(HWPFStream tableStream, String[] entries)
        {
            byte[] header = new byte[6];
            LittleEndian.PutShort(header, 0, unchecked((short)0xffff));

            if (entries == null || entries.Length == 0)
            {
                LittleEndian.PutInt(header, 2, 0);
                tableStream.Write(header);
                return 6;
            }

            LittleEndian.PutInt(header, 2, entries.Length);
            tableStream.Write(header);
            int size = 6;

            foreach (String entry in entries)
            {
                byte[] buf = new byte[entry.Length * 2 + 2];
                LittleEndian.PutShort(buf, 0, (short)entry.Length);
                StringUtil.PutUnicodeLE(entry, buf, 2);
                tableStream.Write(buf);
                size += buf.Length;
            }

            return size;
        }

    }
}


