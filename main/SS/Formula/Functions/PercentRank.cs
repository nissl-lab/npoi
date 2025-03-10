using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    /// <summary>
    /// Implementation of 'Analysis Toolpak' the Excel function PERCENTRANK()
    /// </summary>
    public class PercentRank : Function
    {
        public static Function instance = new PercentRank();
        private PercentRank()
        {
            // Enforce singleton
        }
        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length < 2)
            {
                return ErrorEval.VALUE_INVALID;
            }
            double x;
            try
            {
                ValueEval ev = OperandResolver.GetSingleValue(args[1], srcRowIndex, srcColumnIndex);
                x = OperandResolver.CoerceValueToDouble(ev);
            }
            catch (EvaluationException e)
            {
                ValueEval error = e.GetErrorEval();
                if (error == ErrorEval.NUM_ERROR)
                {
                    return error;
                }
                return ErrorEval.NUM_ERROR;
            }

            List<Double> numbers = new List<double>();
            try
            {
                List<ValueEval> values = getValues(args[0], srcRowIndex, srcColumnIndex);
                foreach (ValueEval ev in values)
                {
                    if (ev is BlankEval || ev is MissingArgEval)
                    {
                        //skip
                    }
                    else
                    {
                        numbers.Add(OperandResolver.CoerceValueToDouble(ev));
                    }
                }
            }
            catch (EvaluationException e)
            {
                ValueEval error = e.GetErrorEval();
                if (error != ErrorEval.NA)
                {
                    return error;
                }
                return ErrorEval.NUM_ERROR;
            }

            if (numbers.Count==0)
            {
                return ErrorEval.NUM_ERROR;
            }

            int significance = 3;
            if (args.Length > 2)
            {
                try
                {
                    ValueEval ev = OperandResolver.GetSingleValue(args[2], srcRowIndex, srcColumnIndex);
                    significance = OperandResolver.CoerceValueToInt(ev);
                }
                catch (EvaluationException e)
                {
                    return e.GetErrorEval();
                }
            }

            return calculateRank(numbers, x, significance, true);
        }
        private ValueEval calculateRank(List<Double> numbers, double x, int significance, bool recurse)
        {
            double closestMatchBelow = Double.MinValue;
            double closestMatchAbove = Double.MaxValue;
            if (recurse)
            {
                foreach (Double d in numbers)
                {
                    if (d <= x && d > closestMatchBelow) closestMatchBelow = d;
                    if (d > x && d < closestMatchAbove) closestMatchAbove = d;
                }
            }
            if (!recurse || closestMatchBelow == x || closestMatchAbove == x)
            {
                int lessThanCount = 0;
                int greaterThanCount = 0;
                foreach (Double d in numbers)
                {
                    if (d < x) lessThanCount++;
                    else if (d > x) greaterThanCount++;
                }
                if (greaterThanCount == numbers.Count|| lessThanCount == numbers.Count)
                {
                    return ErrorEval.NA;
                }
                var result = (double)lessThanCount / (double)(lessThanCount + greaterThanCount);
                return new NumberEval(Math.Floor(result*Math.Pow(10,significance))/Math.Pow(10, significance));
            }
            else
            {
                ValueEval belowRank = calculateRank(numbers, closestMatchBelow, significance, false);
                if (belowRank is not NumberEval below) {
                    return belowRank;
                }
                ValueEval aboveRank = calculateRank(numbers, closestMatchAbove, significance, false);
                if (aboveRank is not NumberEval above) {
                    return aboveRank;
                }

                double diff = closestMatchAbove - closestMatchBelow;
                double pos = x - closestMatchBelow;
                double rankDiff = above.NumberValue - below.NumberValue;
                var result = below.NumberValue + (rankDiff * (pos / diff));
                return new NumberEval(Math.Round(result,significance));
            }
        }
        private List<ValueEval> getValues(ValueEval eval, int srcRowIndex, int srcColumnIndex)
        {
            if (eval is AreaEval ae)
            {
                List<ValueEval> list = new List<ValueEval>();
                for (int r = ae.FirstRow; r <= ae.LastRow; r++)
                {
                    for (int c = ae.FirstColumn; c <= ae.LastColumn; c++)
                    {
                        list.Add(OperandResolver.GetSingleValue(ae.GetAbsoluteValue(r, c), r, c));
                    }
                }
                return list;
            }
            else
            {
                return new List<ValueEval>() { OperandResolver.GetSingleValue(eval, srcRowIndex, srcColumnIndex) };
            }
        }
    }
}
