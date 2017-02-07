using NPOI.SS.UserModel;

namespace NPOI.XSSF.Streaming.Values
{
    public class NumericFormulaValue : FormulaValue
    {
        public double PreEvaluatedValue;

        public override CellType GetFormulaType()
        {
            return CellType.Numeric;
        }
    }
}
