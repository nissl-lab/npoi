using NPOI.SS.UserModel;

namespace NPOI.XSSF.Streaming.Values
{
    public abstract class StringValue : Value
    {
        public CellType GetType()
        {
            return CellType.String;
        }
        //We cannot introduce a new type CellType.RICH_TEXT because the types are public so we have to make rich text as a type of string
        public abstract bool IsRichText(); // using the POI style which seems to avoid "instanceof".
    }
}
