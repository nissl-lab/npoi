using NPOI.SS.Formula.Eval;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    public class Frequency : Fixed2ArgFunction
    {
        public static Function instance = new Frequency();

        private Frequency()
        {
            // enforce singleton
        }
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            MatrixFunction.MutableValueCollector collector = new MatrixFunction.MutableValueCollector(false, false);

            double[] values;
            double[] bins;
            try
            {
                values = collector.collectValues(arg0);
                bins = collector.collectValues(arg1);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }

            // can bins be not sorted?
            //bins = Arrays.stream(bins).sorted().distinct().toArray();

            int[] histogram = Histogram(values, bins);
            NumberEval[] result = Arrays.stream(histogram).boxed().map(NumberEval::new).toArray(NumberEval[]::new);
            return new CacheAreaEval(srcRowIndex, srcColumnIndex,
                    srcRowIndex + result.Length - 1, srcColumnIndex, result);
        }

        static int FindBin(double value, double[] bins)
        {

            int idx = Arrays.binarySearch(bins, value);
            return idx >= 0 ? idx + 1 : -idx;
        }

        static int[] Histogram(double[] values, double[] bins)
        {
            int[] histogram = new int[bins.Length + 1];
            foreach (double val in values)
            {
                histogram[FindBin(val, bins) - 1]++;
            }
            return histogram;
        }
    }
}
