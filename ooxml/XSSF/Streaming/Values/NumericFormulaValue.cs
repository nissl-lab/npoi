using NPOI.SS.UserModel;

namespace NPOI.XSSF.Streaming.Values
{
    public class NumericFormulaValue : FormulaValue
    {
        public double _preEvaluatedValue;

        public override CellType GetFormulaType()
        {
            return CellType.Numeric;
        }
    }
}
