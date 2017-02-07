using NPOI.SS.UserModel;

namespace NPOI.XSSF.Streaming.Values
{
    public abstract class FormulaValue : Value
    {
        public string Value;
        public CellType GetType()
        {
            return CellType.Formula;
        }

        public abstract CellType GetFormulaType();
    }
}
