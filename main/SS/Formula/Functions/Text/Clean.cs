using System;
using System.Text;
using NPOI.SS.Formula.Eval;

namespace NPOI.SS.Formula.Functions
{
    public class Clean : SingleArgTextFunc
    {

        public override ValueEval Evaluate(String arg)
        {
            StringBuilder result = new StringBuilder();
            for (int i = 0; i < arg.Length; i++)
            {
                char c = arg[i];
                if (TextFunction.IsPrintable(c))
                {
                    result.Append(c);
                }
            }
            return new StringEval(result.ToString());
        }
    }
}
