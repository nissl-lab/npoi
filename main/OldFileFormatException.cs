using System;

namespace NPOI
{
    [Serializable]
    public class OldFileFormatException : UnsupportedFileFormatException
    {
        public OldFileFormatException(String s)
            : base(s)
        { }

    }
}
