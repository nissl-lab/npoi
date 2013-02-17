namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula.Eval;
    using System;

    public class Poisson : Fixed3ArgFunction
    {
        private const double DEFAULT_RETURN_RESULT = 1;
        /**
         * This checks is x = 0 and the mean = 0.
         * Excel currently returns the value 1 where as the
         * maths common implementation will error.
         * @param x  The number.
         * @param mean The mean.
         * @return If a default value should be returned.
         */
        private bool IsDefaultResult(double x, double mean)
        {

            if (x == 0 && mean == 0)
            {
                return true;
            }
            return false;
        }

        private bool CheckArgument(double aDouble)
        {

            NumericFunction.CheckValue(aDouble);

            // make sure that the number is positive
            if (aDouble < 0)
            {
                throw new EvaluationException(ErrorEval.NUM_ERROR);
            }

            return true;
        }

        private double probability(int k, double lambda)
        {
            return Math.Pow(lambda, k) * Math.Exp(-lambda) / Factorial(k);
        }

        private double cumulativeProbability(int x, double lambda)
        {
            double result = 0;
            for (int k = 0; k <= x; k++)
            {
                result += probability(k, lambda);
            }
            return result;
        }

        /** All long-representable factorials */
        private long[] FACTORIALS = new long[] {
                           1L,                  1L,                   2L,
                           6L,                 24L,                 120L,
                         720L,               5040L,               40320L,
                      362880L,            3628800L,            39916800L,
                   479001600L,         6227020800L,         87178291200L,
               1307674368000L,     20922789888000L,     355687428096000L,
            6402373705728000L, 121645100408832000L, 2432902008176640000L };


        public long Factorial(int n)
        {
            if (n < 0 || n > 20)
            {
                throw new ArgumentException("Valid argument should be in the range [0..20]");
            }
            return FACTORIALS[n];
        }
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1, ValueEval arg2)
        {
            // arguments/result for this function
            double mean = 0;
            double x = 0;
            bool cumulative = ((BoolEval)arg2).BooleanValue;
            double result = 0;

            try
            {
                x = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                mean = NumericFunction.SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);

                // check for default result : excel implementation for 0,0
                // is different to Math Common.
                if (IsDefaultResult(x, mean))
                {
                    return new NumberEval(DEFAULT_RETURN_RESULT);
                }
                // check the arguments : as per excel function def
                CheckArgument(x);
                CheckArgument(mean);

                // truncate x : as per excel function def
                if (cumulative)
                {
                    result = cumulativeProbability((int)x, mean);
                }
                else
                {
                    result = probability((int)x, mean);
                }

                // check the result
                NumericFunction.CheckValue(result);

            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }

            return new NumberEval(result);
        }
    }
}