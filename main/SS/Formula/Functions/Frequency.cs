/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.Util;
    using NPOI.Util.ArrayExtensions;


    /// <summary>
    /// <para>
    /// Implementation of Excel 'Analysis ToolPak' function FREQUENCY()<br>
    /// Returns a frequency distribution as a vertical array
    /// </para>
    /// <para>
    /// <b>Syntax</b><br>
    /// <b>FREQUENCY</b>(<b>data_array</b>, <b>bins_array</b>)
    /// </para>
    /// <para>
    /// <b>data_array</b> Required. An array of or reference to a Set of values for which you want to count frequencies.
    /// If data_array contains no values, FREQUENCY returns an array of zeros.<br>
    /// <b>bins_array</b> Required. An array of or reference to intervals into which you want to group the values in data_array.
    /// If bins_array contains no values, FREQUENCY returns the number of elements in data_array.<br>
    /// </para>
    /// </summary>
    public class Frequency : Fixed2ArgFunction
    {
        public static  Function Instance = new Frequency();

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
            catch(EvaluationException e)
            {
                return e.GetErrorEval();
            }

            // can bins be not sorted?
            //bins = Arrays.stream(bins).sorted().distinct().ToArray();

            int[] histogram = Histogram(values, bins);

            NumberEval[] result = new NumberEval[histogram.Length]; // Arrays.Stream(histogram).boxed().map(NumberEval::new).ToArray(NumberEval[]::new);
            for(var i = 0; i<histogram.Length; i++)
            {
                result[i] = new NumberEval(histogram[i]);
            }
            return new CacheAreaEval(srcRowIndex, srcColumnIndex,
                    srcRowIndex + result.Length - 1, srcColumnIndex, result);
        }

        public static int FindBin(double value, double[] bins)
        {
            int idx = Array.BinarySearch(bins, value);
            return idx >= 0 ? idx + 1 : -idx;
        }

        public static int[] Histogram(double[] values, double[] bins)
        {
            int[] histogram = new int[bins.Length + 1];
            foreach(double val in values)
            {
                histogram[FindBin(val, bins) - 1]++;
            }
            return histogram;
        }
    }
}

