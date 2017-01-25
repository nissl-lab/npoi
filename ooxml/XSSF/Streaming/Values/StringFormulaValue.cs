using NPOI.SS.UserModel;

namespace NPOI.XSSF.Streaming.Values
{
    public class StringFormulaValue : FormulaValue
    {
        public string _preEvaluatedValue;

        public override CellType GetFormulaType()
        {
            return CellType.String;
        }

    }
}
