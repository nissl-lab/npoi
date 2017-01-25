using NPOI.SS.UserModel;

namespace NPOI.XSSF.Streaming.Values
{
    public class BooleanValue : Value
    {
        public bool _value;
        public CellType GetType()
        {
            return CellType.Boolean;
        }

    }
}
