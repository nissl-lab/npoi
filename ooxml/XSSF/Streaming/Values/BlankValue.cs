using NPOI.SS.UserModel;

namespace NPOI.XSSF.Streaming.Values
{
    public class BlankValue : Value
    {
        CellType Value.GetType()
        {
            return CellType.Blank;
        }
    }
}
