using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI
{
    public class OldFileFormatException : Exception
    {
        public OldFileFormatException(String s)
            : base(s)
        { }

    }
}
