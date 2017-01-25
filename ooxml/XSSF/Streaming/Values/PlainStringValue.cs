namespace NPOI.XSSF.Streaming.Values
{
    public class PlainStringValue : StringValue
    {
        public string _value;

        public override bool IsRichText()
        {
            return false;
        }
    }
}
