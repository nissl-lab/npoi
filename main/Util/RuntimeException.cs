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

    //public class InvalidOperationException : RuntimeException
    //{
    //    public InvalidOperationException()
    //        : base()
    //    {

    //    }
    //    public InvalidOperationException(string message)
    //        : base(message)
    //    {
    //    }
    //    public InvalidOperationException(Exception e)
    //        : base("", e)
    //    {
    //    }
    //    public InvalidOperationException(string exception, Exception ex)
    //        : base(exception, ex)
    //    {

    //    }
    //}
}
