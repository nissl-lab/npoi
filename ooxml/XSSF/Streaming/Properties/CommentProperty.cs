namespace NPOI.XSSF.Streaming.Properties
{
    public class CommentProperty : Property
    {
        public CommentProperty(object value) : base(value)
        {

        }

        public override int GetType()
        {
            return COMMENT;
        }
    }
}
