using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.Util
{
    internal class AssertFailedException : ApplicationException
    {
        public AssertFailedException(string message)
            : base(message)
        {

        }
    }
}
