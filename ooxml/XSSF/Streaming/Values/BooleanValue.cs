using NPOI.SS.UserModel;

namespace NPOI.XSSF.Streaming.Values
{
    public class BooleanValue : Value
    {
        public bool Value { get; set; }
        public CellType GetType()
        {
            return CellType.Boolean;
        }

    }
}
