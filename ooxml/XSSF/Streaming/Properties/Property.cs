namespace NPOI.XSSF.Streaming.Properties
{
    public abstract class Property
    {
        public const int COMMENT = 1;
        public const int HYPERLINK = 2;
        public object _value;
        public Property _next;

        protected Property(object value)
        {
            _value = value;
        }
        public abstract int GetType();
    }
}
