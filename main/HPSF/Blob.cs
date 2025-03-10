using NPOI.Util;

namespace NPOI.HPSF
{
    internal sealed class Blob
    {
        private readonly byte[] _value;

        public Blob(byte[] data, int offset)
        {
            int size = LittleEndian.GetInt(data, offset);

            if (size == 0)
            {
                _value = System.Array.Empty<byte>();
                return;
            }

            _value = LittleEndian.GetByteArray(data, offset
                    + LittleEndian.INT_SIZE, size);
        }

        public int Size
        {
            get { return LittleEndian.INT_SIZE + _value.Length; }
        }
    }
}