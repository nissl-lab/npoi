using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI
{
    public class EncryptedDocumentException : InvalidOperationException
    {
        public EncryptedDocumentException(String s)
            : base(s)
        { }

    }
}
