/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/
/*
 * Created on May 30, 2005
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using NPOI.Util;
    using System;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *
     * Library for common statistics functions
     */
    public class StatsLib
    {

        private StatsLib() { }


        /**
         * returns the mean of deviations from mean.
         * @param v
         */
        public static double avedev(double[] v)
        {
            double r = 0;
            double m = 0;
            double s = 0;
            for (int i = 0, iSize = v.Length; i < iSize; i++)
            {
                s += v[i];
            }
            m = s / v.Length;
            s = 0;
            for (int i = 0, iSize = v.Length; i < iSize; i++)
            {
                s += Math.Abs(v[i] - m);
            }
            r = s / v.Length;
            return r;
        }

        public static double stdev(double[] v)
        {
            double r = double.NaN;
            if (v != null && v.Length > 1)
            {
                r = Math.Sqrt(devsq(v) / (v.Length - 1));
            }
            return r;
        }
        public static double var(double[] v)
        {
            double r = Double.NaN;
            if (v != null && v.Length > 1)
            {
                r = devsq(v) / (v.Length - 1);
            }
            return r;
        }

        public static double varp(double[] v)
        {
            double r = Double.NaN;
            if (v != null && v.Length > 1)
            {
                r = devsq(v) / v.Length;
            }
            return r;
        }
        /**
         * if v Is zero Length or Contains no duplicates, return value
         * Is double.NaN. Else returns the value that occurs most times
         * and if there Is a tie, returns the first such value. 
         * @param v
         */
        public static double mode(double[] v)
        {
            double r = double.NaN;

            // very naive impl, may need to be optimized
            if (v != null && v.Length > 1)
            {
                int[] Counts = new int[v.Length];
                Arrays.Fill(Counts, 1);
                for (int i = 0, iSize = v.Length; i < iSize; i++)
                {
                    for (int j = i + 1, jSize = v.Length; j < jSize; j++)
                    {
                        if (v[i] == v[j]) Counts[i]++;
                    }
                }
                double maxv = 0;
                int maxc = 0;
                for (int i = 0, iSize = Counts.Length; i < iSize; i++)
                {
                    if (Counts[i] > maxc)
                    {
                        maxv = v[i];
                        maxc = Counts[i];
                    }
                }
                r = (maxc > 1) ? maxv : double.NaN; // "no-dups" Check
            }
            return r;
        }

        public static double median(double[] v)
        {
            double r = double.NaN;

            if (v != null && v.Length >= 1)
            {
                int n = v.Length;
                Array.Sort(v);
                r = (n % 2 == 0)
                    ? (v[n / 2] + v[n / 2 - 1]) / 2
                    : v[n / 2];
            }

            return r;
        }


        public static double devsq(double[] v)
        {
            double r = double.NaN;
            if (v != null && v.Length >= 1)
            {
                double m = 0;
                double s = 0;
                int n = v.Length;
                for (int i = 0; i < n; i++)
                {
                    s += v[i];
                }
                m = s / n;
                s = 0;
                for (int i = 0; i < n; i++)
                {
                    s += (v[i] - m) * (v[i] - m);
                }

                r = (n == 1)
                        ? 0
                        : s;
            }
            return r;
        }

        /*
         * returns the kth largest element in the array. Duplicates
         * are considered as distinct values. Hence, eg.
         * for array {1,2,4,3,3} & k=2, returned value Is 3.
         * <br/>
         * k <= 0 & k >= v.Length and null or empty arrays
         * will result in return value double.NaN
         * @param v
         * @param k
         */
        public static double kthLargest(double[] v, int k)
        {
            double r = double.NaN;
            k--; // since arrays are 0-based
            if (v != null && v.Length > k && k >= 0)
            {
                Array.Sort(v);
                r = v[v.Length - k - 1];
            }
            return r;
        }

        /*
         * returns the kth smallest element in the array. Duplicates
         * are considered as distinct values. Hence, eg.
         * for array {1,1,2,4,3,3} & k=2, returned value Is 1.
         * <br/>
         * k <= 0 & k >= v.Length or null array or empty array
         * will result in return value double.NaN
         * @param v
         * @param k
         */
        public static double kthSmallest(double[] v, int k)
        {
            double r = double.NaN;
            k--; // since arrays are 0-based
            if (v != null && v.Length > k && k >= 0)
            {
                Array.Sort(v);
                r = v[k];
            }
            return r;
        }
    }
}