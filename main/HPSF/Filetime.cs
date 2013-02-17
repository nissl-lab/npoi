using System.IO;
using NPOI.Util;

namespace NPOI.HPSF
{
    public class Filetime
    {
        public const int SIZE = LittleEndian.INT_SIZE * 2;

        private int _dwHighDateTime;
        private int _dwLowDateTime;

        public Filetime(byte[] data, int offset)
        {
            _dwLowDateTime = LittleEndian.GetInt(data, offset + 0
                    * LittleEndian.INT_SIZE);
            _dwHighDateTime = LittleEndian.GetInt(data, offset + 1
                    * LittleEndian.INT_SIZE);
        }

        public Filetime(int low, int high)
        {
            _dwLowDateTime = low;
            _dwHighDateTime = high;
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
            LittleEndian.PutInt(result, 0 * LittleEndian.INT_SIZE, _dwLowDateTime);
            LittleEndian.PutInt(result, 1 * LittleEndian.INT_SIZE, _dwHighDateTime);
            return result;
        }

        public int Write(Stream out1)
        {
            LittleEndian.PutInt(_dwLowDateTime, out1);
            LittleEndian.PutInt(_dwHighDateTime, out1);
            return SIZE;
        }
    }
}