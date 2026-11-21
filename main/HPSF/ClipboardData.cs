using System.IO;
using NPOI.Util;

namespace NPOI.HPSF
{
    public class ClipboardData
    {
        //arbitrarily selected; may need to increase
        private static int MAX_RECORD_LENGTH = 100_000_000;

        private int _format = 0;
        private byte[] _value;
        internal ClipboardData() {}
        internal void Read( LittleEndianByteArrayInputStream lei ) 
        {
            int offset = lei.GetReadIndex();
            int size = lei.ReadInt();

            if ( size < 4 ) {
                //String msg = 
                //    "ClipboardData at offset "+offset+" size less than 4 bytes "+
                //    "(doesn't even have format field!). Setting to format == 0 and hope for the best";
                //LOG.log( POILogger.WARN, msg);
                _format = 0;
                _value = [];
                return;
            }

            _format = lei.ReadInt();
            _value = IOUtils.SafelyAllocate(size - LittleEndianConsts.INT_SIZE, MAX_RECORD_LENGTH);
            lei.ReadFully(_value);
        }

        public byte[] Value
        {
            get { return _value; }
            set { _value = (byte[])value.Clone(); }
        }

        internal byte[] ToByteArray() 
        {
            byte[] result = new byte[LittleEndianConsts.INT_SIZE*2+_value.Length];
            LittleEndianByteArrayOutputStream bos = new LittleEndianByteArrayOutputStream(result,0);
            try 
            {
                bos.WriteInt(LittleEndianConsts.INT_SIZE + _value.Length);
                bos.WriteInt(_format);
                bos.Write(_value);
                return result;
            } 
            finally 
            {
                //IOUtils.CloseQuietly(bos); //bos is not a stream object
            }
        }
    }
}