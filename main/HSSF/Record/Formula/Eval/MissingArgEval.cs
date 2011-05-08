using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSSF.Record.Formula.Eval
{
    public class MissingArgEval : ValueEval
    {
        public static MissingArgEval instance = new MissingArgEval();

        private MissingArgEval()
        {
        }
    }
}
