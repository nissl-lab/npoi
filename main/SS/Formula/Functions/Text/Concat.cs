namespace NPOI.SS.Formula.Functions
{
    using System.Text; 
using Cysharp.Text;
    using NPOI.SS.Formula.Eval;

    public class Concat : FreeRefFunction
    {
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            using var sb = ZString.CreateStringBuilder();

            foreach(ValueEval arg in args)
            {
                try
                {
                    if(arg is AreaEval area)
                    {
                        for(int rn = 0; rn<area.Height; rn++)
                        {
                            for(int cn = 0; cn<area.Width; cn++)
                            {
                                ValueEval ve = area.GetRelativeValue(rn, cn);
                                sb.Append(TextFunction.EvaluateStringArg(ve, ec.RowIndex, ec.ColumnIndex));
                            }
                        }
                    }
                    else
                    {
                        sb.Append(TextFunction.EvaluateStringArg(arg, ec.RowIndex, ec.ColumnIndex));
                    }
                } 
                catch (EvaluationException e) {
                    return e.GetErrorEval();
                }
            }
            return new StringEval(sb.ToString());
        }

    }
}