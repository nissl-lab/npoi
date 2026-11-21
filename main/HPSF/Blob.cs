using NPOI.Util;

namespace NPOI.HPSF
{
    internal sealed class Blob
    {
        //arbitrarily selected; may need to increase
        private static int MAX_RECORD_LENGTH = 1_000_000;
        private byte[] _value;

        internal Blob() { }

        internal void Read(LittleEndianByteArrayInputStream lei)
        {
            int size = lei.ReadInt();
            _value = IOUtils.SafelyAllocate(size, MAX_RECORD_LENGTH);
            if(size > 0)
            {
                lei.ReadFully(_value);
            }
        }
    }
}