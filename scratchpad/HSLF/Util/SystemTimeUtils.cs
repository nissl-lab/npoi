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

namespace NPOI.HSLF.Util
{
    using System;
    using NPOI.Util;

    /**
     * A helper class for dealing with SystemTime Structs, as defined at
     * http://msdn.microsoft.com/library/en-us/sysinfo/base/systemtime_str.asp .
     *
     * Discrepancies between Calendar and SYSTEMTIME:
     *  - that January = 1 in SYSTEMTIME, 0 in Calendar.
     *  - that the day of the week (0) starts on Sunday in SYSTEMTIME, and Monday in Calendar
     * It is also the case that this does not store the timezone, and no... it is not
     * stored as UTC either, but rather the local system time (yuck.)
     *
     * @author Daniel Noll
     * @author Nick Burch
     */
    public class SystemTimeUtils
    {
        /**
         * Get the date found in the byte array, as a java Data object
         */
        public static DateTime GetDate(byte[] data)
        {
            return GetDate(data, 0);
        }
        /**
         * Get the date found in the byte array, as a java Data object
         */
        public static DateTime GetDate(byte[] data, int offset)
        {
            short year=LittleEndian.GetShort(data, offset);
            short month=LittleEndian.GetShort(data, offset + 2);
            // Not actually needed - can be found from day of month
            //cal.Set(Calendar.DAY_OF_WEEK,  LittleEndian.GetShort(data,OffSet+4)+1);
            short day=LittleEndian.GetShort(data, offset + 6);
            short hour=LittleEndian.GetShort(data, offset + 8);
            short minute= LittleEndian.GetShort(data, offset + 10);
            short second= LittleEndian.GetShort(data, offset + 12);
            short millisecond= LittleEndian.GetShort(data, offset + 14);

            return new DateTime(year, month, day, hour, minute, second, millisecond);
        }

        /**
         * Convert the supplied java Date into a SystemTime struct, and write it
         *  into the supplied byte array.
         */
        public static void StoreDate(DateTime date, byte[] dest)
        {
            StoreDate(date, dest, 0);
        }
        /**
         * Convert the supplied java Date into a SystemTime struct, and write it
         *  into the supplied byte array.
         */
        public static void StoreDate(DateTime date, byte[] dest, int offset)
        {
            LittleEndian.PutShort(dest, offset + 0, (short)date.Year);
            LittleEndian.PutShort(dest, offset + 2, (short)date.Month);
            LittleEndian.PutShort(dest, offset + 4, (short)date.DayOfWeek);
            LittleEndian.PutShort(dest, offset + 6, (short)date.Day);
            LittleEndian.PutShort(dest, offset + 8, (short)date.Hour);
            LittleEndian.PutShort(dest, offset + 10, (short)date.Minute);
            LittleEndian.PutShort(dest, offset + 12, (short)date.Second);
            LittleEndian.PutShort(dest, offset + 14, (short)date.Millisecond);
        }
    }
}



