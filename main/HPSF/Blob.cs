using NPOI.Util;

namespace NPOI.HPSF
{
    internal sealed class Blob
    {
        private byte[] _value;

        internal Blob() { }

        internal void Read(LittleEndianByteArrayInputStream lei)
        {
            int size = lei.ReadInt();
            _value = new byte[size];
            if(size > 0)
            {
                lei.ReadFully(_value);
            }
        }
    }
}