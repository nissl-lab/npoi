using NPOI.SS.UserModel;

namespace NPOI.XSSF.Streaming.Values
{
    public class StringFormulaValue : FormulaValue
    {
        public string PreEvaluatedValue;

        public override CellType GetFormulaType()
        {
            return CellType.String;
        }

    }
}
