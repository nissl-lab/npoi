namespace NPOI.XSSF.Streaming.Properties
{
    public class HyperlinkProperty : Property
    {
        public HyperlinkProperty(object value) : base(value)
        {

        }

        public override int GetType()
        {
            return HYPERLINK;
        }
    }
}
