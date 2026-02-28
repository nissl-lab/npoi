using System;
using System.IO;
using NPOI.Util;

namespace NPOI.HPSF
{
    public class Filetime
    {
        /**
     * The difference between the Windows epoch (1601-01-01
     * 00:00:00) and the Unix epoch (1970-01-01 00:00:00) in
     * milliseconds.
     */
        private static long EPOCH_DIFF = -11644473600000L;
        private const int SIZE = LittleEndianConsts.INT_SIZE * 2;
        private static long UINT_MASK = 0x00000000FFFFFFFFL;
        private static long NANO_100 = 1000L * 10L;
        private int _dwHighDateTime;
        private int _dwLowDateTime;

        internal Filetime(DateTime date)
        {
            long filetime = date.ToFileTime();
            _dwHighDateTime = (int) ((filetime >>> 32) & UINT_MASK);
            _dwLowDateTime = (int) (filetime & UINT_MASK);
        }
        internal Filetime() { }

        internal Filetime(int low, int high)
        {
            _dwLowDateTime = low;
            _dwHighDateTime = high;
        }

        internal void Read(LittleEndianByteArrayInputStream lei)
        {
            _dwLowDateTime = lei.ReadInt();
            _dwHighDateTime = lei.ReadInt();
        }
        public long High
        {
            get { return _dwHighDateTime; }
        }

        public long Low
        {
            get { return _dwLowDateTime; }
        }

        public byte[] ToByteArray()
        {
            byte[] result = new byte[SIZE];
            LittleEndian.PutInt(result, 0 * LittleEndianConsts.INT_SIZE, _dwLowDateTime);
            LittleEndian.PutInt(result, 1 * LittleEndianConsts.INT_SIZE, _dwHighDateTime);
            return result;
        }

        public int Write(Stream out1)
        {
            LittleEndian.PutInt(_dwLowDateTime, out1);
            LittleEndian.PutInt(_dwHighDateTime, out1);
            return SIZE;
        }

        internal DateTime GetJavaValue()
        {
            long l = (((long)_dwHighDateTime) << 32) | (_dwLowDateTime & UINT_MASK);
            return DateTime.FromFileTime(l);
        }

        /**
         * Converts a Windows FILETIME into a {@link Date}. The Windows
         * FILETIME structure holds a date and time associated with a
         * file. The structure identifies a 64-bit integer specifying the
         * number of 100-nanosecond intervals which have passed since
         * January 1, 1601.
         *
         * @param filetime The filetime to convert.
         * @return The Windows FILETIME as a {@link Date}.
         */
        public static DateTime FiletimeToDate(long filetime)
        {
            //long ms_since_16010101 = filetime / NANO_100;
            //long ms_since_19700101 = ms_since_16010101 + EPOCH_DIFF;
            //return new DateTime(ms_since_19700101);
            return DateTime.FromFileTime(filetime);
        }

        /**
         * Converts a {@link Date} into a filetime.
         *
         * @param date The date to be converted
         * @return The filetime
         *
         * @see #filetimeToDate(long)
         */
        public static long DateToFileTime(DateTime date)
        {
            //long ms_since_19700101 = date.getTime();
            //long ms_since_16010101 = ms_since_19700101 - EPOCH_DIFF;
            //return ms_since_16010101 * NANO_100;
            return date.ToFileTime();
        }

        /**
         * Return {@code true} if the date is undefined
         *
         * @param date the date
         * @return {@code true} if the date is undefined
         */
        public static bool IsUndefined(DateTime? date)
        {
            return !date.HasValue || DateToFileTime(date.Value) == 0;
        }
    }
}