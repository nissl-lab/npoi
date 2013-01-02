using System.IO;
using NPOI.Util;

namespace NPOI.HPSF
{
    public class ClipboardData
    {
        //private static final POILogger logger = POILogFactory
        //    .getLogger( ClipboardData.class );

        private int _format;
        private byte[] _value;

        public ClipboardData(byte[] data, int offset)
        {
            int size = LittleEndian.GetInt(data, offset);

            if (size < 4)
            {
                //logger.log( POILogger.WARN, "ClipboardData at offset ",
                //        Integer.valueOf( offset ), " size less than 4 bytes "
                //                + "(doesn't even have format field!). "
                //                + "Setting to format == 0 and hope for the best" );
                _format = 0;
                _value = new byte[0];
                return;
            }

            _format = LittleEndian.GetInt(data, offset + LittleEndian.INT_SIZE);
            _value = LittleEndian.GetByteArray(data, offset
                    + LittleEndian.INT_SIZE * 2, size - LittleEndian.INT_SIZE);
        }

        public int Size
        {
            get { return LittleEndian.INT_SIZE*2 + _value.Length; }
        }

        public byte[] Value
        {
            get { return _value; }
        }

        public byte[] ToByteArray()
        {
            byte[] result = new byte[Size];
            LittleEndian.PutInt(result, 0 * LittleEndian.INT_SIZE,
                    LittleEndian.INT_SIZE + _value.Length);
            LittleEndian.PutInt(result, 1 * LittleEndian.INT_SIZE, _format);
            System.Array.Copy(_value, 0, result, LittleEndian.INT_SIZE
                    + LittleEndian.INT_SIZE, _value.Length);
            return result;
        }

        public int Write(Stream out1)
        {
            LittleEndian.PutInt(LittleEndian.INT_SIZE + _value.Length, out1);
            LittleEndian.PutInt(_format, out1);
            out1.Write(_value, 0, _value.Length);
            return 2 * LittleEndian.INT_SIZE + _value.Length;
        }
    }
}