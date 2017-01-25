using NPOI.SS.UserModel;

namespace NPOI.XSSF.Streaming.Values
{
    public class BooleanFormulaValue : FormulaValue
    {
        public bool _preEvaluatedValue;

        public override CellType GetFormulaType()
        {
            return CellType.Boolean;
        }

    }
}
