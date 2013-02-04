using System;

namespace NPOI
{
    [Serializable]
    public class OldFileFormatException : Exception
    {
        public OldFileFormatException(String s)
            : base(s)
        { }

    }
}
