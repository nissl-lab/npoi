using System;
using System.Text;
using NPOI.HSSF.Record.Formula.Eval;
using HSSFFunctions=NPOI.HSSF.Record.Formula.Functions;

namespace TestCases.HSSF.Record.Formula.Eval
{
    internal class EvalInstances
    {
        public static HSSFFunctions.Function Add = TwoOperandNumericOperation.AddEval;
        public static HSSFFunctions.Function Subtract = TwoOperandNumericOperation.SubtractEval;
        public static HSSFFunctions.Function Multiply = TwoOperandNumericOperation.MultiplyEval;
        public static HSSFFunctions.Function Divide = TwoOperandNumericOperation.DivideEval;

        public static HSSFFunctions.Function Power = TwoOperandNumericOperation.PowerEval;

        public static HSSFFunctions.Function Percent = PercentEval.instance;

        public static HSSFFunctions.Function UnaryMinus = UnaryMinusEval.instance;
        public static HSSFFunctions.Function UnaryPlus = UnaryPlusEval.instance;

        public static HSSFFunctions.Function Equal = RelationalOperationEval.EqualEval;
        public static HSSFFunctions.Function LessThan = RelationalOperationEval.LessThanEval;
        public static HSSFFunctions.Function LessEqual = RelationalOperationEval.LessEqualEval;
        public static HSSFFunctions.Function GreaterThan = RelationalOperationEval.GreaterThanEval;
        public static HSSFFunctions.Function GreaterEqual = RelationalOperationEval.GreaterEqualEval;
        public static HSSFFunctions.Function NotEqual = RelationalOperationEval.NotEqualEval;

        public static HSSFFunctions.Function Range = RangeEval.instance;
        public static HSSFFunctions.Function Concat = ConcatEval.instance;

    }
}