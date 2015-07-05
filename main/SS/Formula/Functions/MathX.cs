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
 * Created on May 19, 2005
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using System;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     * This class Is an extension to the standard math library
     * provided by java.lang.Math class. It follows the Math class
     * in that it has a private constructor and all static methods.
     */
    public class MathX
    {


        private MathX() { }


        /**
         * Returns a value rounded to p digits after decimal.
         * If p Is negative, then the number Is rounded to
         * places to the left of the decimal point. eg. 
         * 10.23 rounded to -1 will give: 10. If p Is zero,
         * the returned value Is rounded to the nearest integral
         * value.
         * If n Is negative, the resulting value Is obtained
         * as the round value of absolute value of n multiplied
         * by the sign value of n (@see MathX.sign(double d)). 
         * Thus, -0.6666666 rounded to p=0 will give -1 not 0.
         * If n Is NaN, returned value Is NaN.
         * @param n
         * @param p
         */
        public static double Round(double n, int p)
        {
            double retval;

            if (double.IsNaN(n) || double.IsInfinity(n))
            {
                retval = double.NaN;
            }
            else if (double.MaxValue == n)
                return double.MaxValue;
            else if (double.MinValue == n)
                return 0;
            else
            {
                if (p >= 0)
                {
                    int temp = (int)Math.Pow(10, p);
                    double delta = 0.5;
                    int x = p + 1;
                    while (x > 0)
                    {
                        delta = delta / 10;
                        x--;
                    }
                    retval = (double)(Math.Round((decimal)(n + delta) * temp) / temp);
                }
                else
                {
                    int temp = (int)Math.Pow(10, Math.Abs(p));
                    retval = (double)(Math.Round((decimal)(n) / temp) * temp);
                }
            }

            return retval;
        }

        /**
         * Returns a value rounded-up to p digits after decimal.
         * If p Is negative, then the number Is rounded to
         * places to the left of the decimal point. eg. 
         * 10.23 rounded to -1 will give: 20. If p Is zero,
         * the returned value Is rounded to the nearest integral
         * value.
         * If n Is negative, the resulting value Is obtained
         * as the round-up value of absolute value of n multiplied
         * by the sign value of n (@see MathX.sign(double d)). 
         * Thus, -0.2 rounded-up to p=0 will give -1 not 0.
         * If n Is NaN, returned value Is NaN.
         * @param n
         * @param p
         */
        public static double RoundUp(double n, int p)
        {
            double retval;

            if (double.IsNaN(n) || double.IsInfinity(n))
            {
                retval = double.NaN;
            }
            else if (double.MaxValue == n)
                return double.MaxValue;
            else if (double.MinValue == n)
            {
                double digit = 1;
                while (p > 0)
                {
                    digit = digit / 10;
                    p--;
                }
                return digit;
            }
            else
            {
                if (p != 0)
                {
                    double temp = Math.Pow(10, p);
                    double nat = Math.Abs(n * temp);

                    retval = Sign(n) *
                        ((nat == (long)nat)
                                ? nat / temp
                                : Math.Round(nat + 0.5) / temp);
                }
                else
                {
                    double na = Math.Abs(n);
                    retval = Sign(n) *
                        ((na == (long)na)
                            ? na
                            : (long)na + 1);
                }
            }

            return retval;
        }

        /**
         * Returns a value rounded to p digits after decimal.
         * If p Is negative, then the number Is rounded to
         * places to the left of the decimal point. eg. 
         * 10.23 rounded to -1 will give: 10. If p Is zero,
         * the returned value Is rounded to the nearest integral
         * value.
         * If n Is negative, the resulting value Is obtained
         * as the round-up value of absolute value of n multiplied
         * by the sign value of n (@see MathX.sign(double d)). 
         * Thus, -0.8 rounded-down to p=0 will give 0 not -1.
         * If n Is NaN, returned value Is NaN.
         * @param n
         * @param p
         */
        public static double RoundDown(double n, int p)
        {
            double retval;

            if (double.IsNaN(n) || double.IsInfinity(n))
            {
                retval = double.NaN;
            }
            else if (double.MaxValue == n)
                return double.MaxValue;
            else if (double.MinValue == n)
                return 0;
            else
            {
                if (p != 0)
                {
                    double temp = Math.Pow(10, p);
                    retval = Sign(n) * Math.Round((Math.Abs(n) * temp) - 0.5, MidpointRounding.AwayFromZero) / temp;
                }
                else
                {
                    retval = (long)n;
                }
            }

            return retval;
        }


        /*
         * If d < 0, returns short -1
         * <br/>
         * If d > 0, returns short 1
         * <br/>
         * If d == 0, returns short 0 
         *  If d Is NaN, then 1 will be returned. It Is the responsibility
         * of caller to Check for d IsNaN if some other value Is desired.
         * @param d
         */
        public static short Sign(double d)
        {
            return (short)((d == 0)
                    ? 0
                    : (d < 0)
                            ? -1
                            : 1);
        }

        /**
         * average of all values
         * @param values
         */
        public static double Average(double[] values)
        {
            double ave = 0;
            double sum = 0;
            for (int i = 0, iSize = values.Length; i < iSize; i++)
            {
                sum += values[i];
            }
            ave = sum / values.Length;
            return ave;
        }


        /**
         * sum of all values
         * @param values
         */
        public static double Sum(double[] values)
        {
            double sum = 0;
            for (int i = 0, iSize = values.Length; i < iSize; i++)
            {
                sum += values[i];
            }
            return sum;
        }

        /**
         * sum of squares of all values
         * @param values
         */
        public static double Sumsq(double[] values)
        {
            double sumsq = 0;
            for (int i = 0, iSize = values.Length; i < iSize; i++)
            {
                sumsq += values[i] * values[i];
            }
            return sumsq;
        }


        /**
         * product of all values
         * @param values
         */
        public static double Product(double[] values)
        {
            double product = 0;
            if (values != null && values.Length > 0)
            {
                product = 1;
                for (int i = 0, iSize = values.Length; i < iSize; i++)
                {
                    product *= values[i];
                }
            }
            return product;
        }

        /**
         * min of all values. If supplied array Is zero Length,
         * double.POSITIVE_INFINITY Is returned.
         * @param values
         */
        public static double Min(double[] values)
        {
            double min = double.PositiveInfinity;
            for (int i = 0, iSize = values.Length; i < iSize; i++)
            {
                min = Math.Min(min, values[i]);
            }
            return min;
        }

        /**
         * min of all values. If supplied array Is zero Length,
         * double.NEGATIVE_INFINITY Is returned.
         * @param values
         */
        public static double Max(double[] values)
        {
            double max = double.NegativeInfinity;
            for (int i = 0, iSize = values.Length; i < iSize; i++)
            {
                max = Math.Max(max, values[i]);
            }
            return max;
        }

        /**
         * Note: this function Is different from java.lang.Math.floor(..).
         * 
         * When n and s are "valid" arguments, the returned value Is: Math.floor(n/s) * s;
         * <br/>
         * n and s are invalid if any of following conditions are true:
         * <ul>
         * <li>s Is zero</li>
         * <li>n Is negative and s Is positive</li>
         * <li>n Is positive and s Is negative</li>
         * </ul>
         * In all such cases, double.NaN Is returned.
         * @param n
         * @param s
         */
        public static double Floor(double n, double s)
        {
            double f;

            if ((n < 0 && s > 0) || (n > 0 && s < 0) || (s == 0 && n != 0))
            {
                f = double.NaN;
            }
            else
            {
                f = (n == 0 || s == 0) ? 0 : Math.Floor(n / s) * s;
            }

            return f;
        }

        /**
         * Note: this function Is different from java.lang.Math.ceil(..).
         * 
         * When n and s are "valid" arguments, the returned value Is: Math.ceiling(n/s) * s;
         * <br/>
         * n and s are invalid if any of following conditions are true:
         * <ul>
         * <li>s Is zero</li>
         * <li>n Is negative and s Is positive</li>
         * <li>n Is positive and s Is negative</li>
         * </ul>
         * In all such cases, double.NaN Is returned.
         * @param n
         * @param s
         */
        public static double Ceiling(double n, double s)
        {
            double c;

            if ((n < 0 && s > 0) || (n > 0 && s < 0))
            {
                c = double.NaN;
            }
            else
            {
                c = (n == 0 || s == 0) ? 0 : Math.Ceiling(n / s) * s;
            }

            return c;
        }

        /**
         * <br/> for all n >= 1; factorial n = n * (n-1) * (n-2) * ... * 1 
         * <br/> else if n == 0; factorial n = 1
         * <br/> else if n &lt; 0; factorial n = double.NaN
         * <br/> Loss of precision can occur if n Is large enough.
         * If n Is large so that the resulting value would be greater 
         * than double.MAX_VALUE; double.POSITIVE_INFINITY Is returned.
         * If n &lt; 0, double.NaN Is returned. 
         * @param n
         */
        public static double Factorial(int n)
        {
            double d = 1;

            if (n >= 0)
            {
                if (n <= 170)
                {
                    for (int i = 1; i <= n; i++)
                    {
                        d *= i;
                    }
                }
                else
                {
                    d = double.PositiveInfinity;
                }
            }
            else
            {
                d = double.NaN;
            }
            return d;
        }


        /**
         * returns the remainder resulting from operation:
         * n / d. 
         * <br/> The result has the sign of the divisor.
         * <br/> Examples:
         * <ul>
         * <li>mod(3.4, 2) = 1.4</li>
         * <li>mod(-3.4, 2) = 0.6</li>
         * <li>mod(-3.4, -2) = -1.4</li>
         * <li>mod(3.4, -2) = -0.6</li>
         * </ul>
         * If d == 0, result Is NaN
         * @param n
         * @param d
         */
        public static double Mod(double n, double d)
        {
            double result = 0;

            if (d == 0)
            {
                result = double.NaN;
            }
            else if (Sign(n) == Sign(d))
            {
                //double t = Math.Abs(n / d);
                //t = t - (long)t;
                //result = sign(d) * Math.Abs(t * d);
                result = n % d;
            }
            else
            {
                //double t = Math.Abs(n / d);
                //t = t - (long)t;
                //t = Math.Ceiling(t) - t;
                //result = sign(d) * Math.Abs(t * d);
                result = ((n % d) + d) % d;
            }

            return result;
        }


        /**
         * inverse hyperbolic cosine
         * @param d
         */
        public static double Acosh(double d)
        {
            return Math.Log(Math.Sqrt(Math.Pow(d, 2) - 1) + d);
        }

        /**
         * inverse hyperbolic sine
         * @param d
         */
        public static double Asinh(double d)
        {
            double d2 = d * d;
            return Math.Log(Math.Sqrt(d * d + 1) + d);
        }

        /**
         * inverse hyperbolic tangent
         * @param d
         */
        public static double Atanh(double d)
        {
            return Math.Log((1 + d) / (1 - d)) / 2;
        }

        /**
         * hyperbolic cosine
         * @param d
         */
        public static double Cosh(double d)
        {
            double ePowX = Math.Pow(Math.E, d);
            double ePowNegX = Math.Pow(Math.E, -d);
            d = (ePowX + ePowNegX) / 2;
            return d;
        }

        /**
         * hyperbolic sine
         * @param d
         */
        public static double Sinh(double d)
        {
            double ePowX = Math.Pow(Math.E, d);
            double ePowNegX = Math.Pow(Math.E, -d);
            d = (ePowX - ePowNegX) / 2;
            return d;
        }

        /**
         * hyperbolic tangent
         * @param d
         */
        public static double Tanh(double d)
        {
            double ePowX = Math.Pow(Math.E, d);
            double ePowNegX = Math.Pow(Math.E, -d);
            d = (ePowX - ePowNegX) / (ePowX + ePowNegX);
            return d;
        }

        /**
         * returns the sum of product of corresponding double value in each
         * subarray. It Is the responsibility of the caller to Ensure that
         * all the subarrays are of equal Length. If the subarrays are
         * not of equal Length, the return value can be Unpredictable.
         * @param arrays
         */
        public static double SumProduct(double[][] arrays)
        {
            double d = 0;

            try
            {
                int narr = arrays.Length;
                int arrlen = arrays[0].Length;

                for (int j = 0; j < arrlen; j++)
                {
                    double t = 1;
                    for (int i = 0; i < narr; i++)
                    {
                        t *= arrays[i][j];
                    }
                    d += t;
                }
            }
            catch (IndexOutOfRangeException)
            {
                d = double.NaN;
            }

            return d;
        }

        /**
         * returns the sum of difference of squares of corresponding double 
         * value in each subarray: ie. sigma (xarr[i]^2-yarr[i]^2) 
         * <br/>
         * It Is the responsibility of the caller 
         * to Ensure that the two subarrays are of equal Length. If the 
         * subarrays are not of equal Length, the return value can be 
         * Unpredictable.
         * @param xarr
         * @param yarr
         */
        public static double Sumx2my2(double[] xarr, double[] yarr)
        {
            double d = 0;

            try
            {
                for (int i = 0, iSize = xarr.Length; i < iSize; i++)
                {
                    d += (xarr[i] + yarr[i]) * (xarr[i] - yarr[i]);
                }
            }
            catch (IndexOutOfRangeException)
            {
                d = double.NaN;
            }

            return d;
        }

        /**
         * returns the sum of sum of squares of corresponding double 
         * value in each subarray: ie. sigma (xarr[i]^2 + yarr[i]^2) 
         * <br/>
         * It Is the responsibility of the caller 
         * to Ensure that the two subarrays are of equal Length. If the 
         * subarrays are not of equal Length, the return value can be 
         * Unpredictable.
         * @param xarr
         * @param yarr
         */
        public static double Sumx2py2(double[] xarr, double[] yarr)
        {
            double d = 0;

            try
            {
                for (int i = 0, iSize = xarr.Length; i < iSize; i++)
                {
                    d += (xarr[i] * xarr[i]) + (yarr[i] * yarr[i]);
                }
            }
            catch (IndexOutOfRangeException )
            {
                d = double.NaN;
            }

            return d;
        }


        /**
         * returns the sum of squares of difference of corresponding double 
         * value in each subarray: ie. sigma ( (xarr[i]-yarr[i])^2 ) 
         * <br/>
         * It Is the responsibility of the caller 
         * to Ensure that the two subarrays are of equal Length. If the 
         * subarrays are not of equal Length, the return value can be 
         * Unpredictable.
         * @param xarr
         * @param yarr
         */
        public static double Sumxmy2(double[] xarr, double[] yarr)
        {
            double d = 0;

            try
            {
                for (int i = 0, iSize = xarr.Length; i < iSize; i++)
                {
                    double t = (xarr[i] - yarr[i]);
                    d += t * t;
                }
            }
            catch (IndexOutOfRangeException )
            {
                d = double.NaN;
            }

            return d;
        }

        /**
         * returns the total number of combinations possible when
         * k items are chosen out of total of n items. If the number
         * Is too large, loss of precision may occur (since returned
         * value Is double). If the returned value Is larger than
         * double.MAX_VALUE, double.POSITIVE_INFINITY Is returned.
         * If either of the parameters Is negative, double.NaN Is returned.
         * @param n
         * @param k
         */
        public static double NChooseK(int n, int k)
        {
            double d = 1;
            if (n < 0 || k < 0 || n < k)
            {
                d = double.NaN;
            }
            else
            {
                int minnk = Math.Min(n - k, k);
                int maxnk = Math.Max(n - k, k);
                for (int i = maxnk; i < n; i++)
                {
                    d *= i + 1;
                }
                d /= Factorial(minnk);
            }

            return d;
        }

    }
}