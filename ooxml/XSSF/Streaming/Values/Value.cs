using NPOI.SS.UserModel;

namespace NPOI.XSSF.Streaming.Values
{
    public interface Value
    {
        CellType GetType();
    }
}
