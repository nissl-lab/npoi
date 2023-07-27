using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.Formula.Functions
{
    public sealed class FloorMath : FloorCeilingMathBase
    {
        private FloorMath()
        {

        }
        public static readonly FloorMath Instance = new();

        protected override double EvaluateMajorDirection(double number)
            => Math.Floor(number);

        protected override double EvaluateAlternativeDirection(double number)
            => Math.Ceiling(number);
    }
}
