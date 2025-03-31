using System;
using System.Collections.Generic;
using System.Text; 
using Cysharp.Text;

namespace NPOI.Util
{
    internal sealed class AssertFailedException : ApplicationException
    {
        public AssertFailedException(string message)
            : base(message)
        {

        }
    }
}
