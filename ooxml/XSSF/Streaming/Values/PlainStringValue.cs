namespace NPOI.XSSF.Streaming.Values
{
    public class PlainStringValue : StringValue
    {
        public string Value;

        public override bool IsRichText()
        {
            return false;
        }
    }
}
