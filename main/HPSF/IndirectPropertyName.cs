namespace NPOI.HPSF
{
    public class IndirectPropertyName
    {
        private readonly CodePageString _value;

        public IndirectPropertyName(byte[] data, int offset)
        {
            _value = new CodePageString(data, offset);
        }

        public int Size
        {
            get { return _value.Size; }
        } 
    }
}