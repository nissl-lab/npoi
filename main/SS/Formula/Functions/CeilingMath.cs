using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.Formula.Functions
{
    public sealed class CeilingMath : FloorCeilingMathBase
    {
        private CeilingMath()
        {
        }
        public static readonly CeilingMath Instance = new();
        protected override double EvaluateMajorDirection(double number)
            => Math.Ceiling(number);

        protected override double EvaluateAlternativeDirection(double number)
            => Math.Floor(number);
    }
}
