using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    public class MatrixFunction
    {
        public class MutableValueCollector : MultiOperandNumericFunction
        {
            public MutableValueCollector(bool isReferenceBoolCounted, bool isBlankCounted) : base(isReferenceBoolCounted, isBlankCounted)
            {

            }
            public double[] collectValues(params ValueEval[] operands)
            {
                return GetNumberArray(operands);
            }
            protected internal override double Evaluate(double[] values)
            {
                throw new InvalidOperationException("should not be called");
            }
        }
    }
}
