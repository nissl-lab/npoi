using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.Util
{
    public sealed class AssertFailedException : ApplicationException
    {
        public AssertFailedException(string message)
            : base(message)
        {

        }
    }
}
