using System;

namespace NPOI
{
    [Serializable]
    public class EncryptedDocumentException : InvalidOperationException
    {
        public EncryptedDocumentException(string message)
            : base(message)
        { }
        public EncryptedDocumentException(string message, Exception cause)
            : base(message, cause)
        { }
        public EncryptedDocumentException(Exception cause)
            : base(cause.Message)
        {
        }
    }
}
