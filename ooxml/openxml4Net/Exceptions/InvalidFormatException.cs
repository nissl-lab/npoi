using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXml4Net.Exceptions
{
    public class InvalidFormatException:OpenXml4NetException
    {
        private string p;
        private InvalidFormatException ex;

        public InvalidFormatException(String message):base(message)
        {
            
        }

        public InvalidFormatException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
