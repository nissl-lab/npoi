using System;

namespace NPOI.Util
{
    [Serializable]
    public class RuntimeException:Exception
    {
        public RuntimeException()
            :base()
        {
            
        }
        public RuntimeException(string message)
            : base(message)
        {
        }
        public RuntimeException(Exception e)
            : base("", e)
        {
        }
        public RuntimeException(string exception, Exception ex)
            : base(exception, ex)
        {

        }
    }

    public class IllegalStateException : RuntimeException
    {
        public IllegalStateException()
            : base()
        {

        }
        public IllegalStateException(string message)
            : base(message)
        {
        }
        public IllegalStateException(Exception e)
            : base("", e)
        {
        }
        public IllegalStateException(string exception, Exception ex)
            : base(exception, ex)
        {

        }
    }
}
