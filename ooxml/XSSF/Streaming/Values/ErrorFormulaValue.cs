using NPOI.SS.UserModel;

namespace NPOI.XSSF.Streaming.Values
{
    public class ErrorFormulaValue : FormulaValue
    {
        public byte _preEvaluatedValue;

        public override CellType GetFormulaType()
        {
            return CellType.Error;
        }

    }
}
