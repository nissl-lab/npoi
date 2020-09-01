using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.Formula.Eval
{
    public class NumberValueArrayEval : ArrayEval
    {
        public double[] NumberValues { get; set; }

        public NumberValueArrayEval(double[] values)
        {
            NumberValues = values;
        }
    }
}
