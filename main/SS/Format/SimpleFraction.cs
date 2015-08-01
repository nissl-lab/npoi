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
*/
using System;
using NPOI.Util;
namespace NPOI.SS.Format
{
    public class SimpleFraction
    {


        /** The denominator. */
        private int denominator;

        /** The numerator. */
        private int numerator;
        /**
         * Create a fraction given a double value and a denominator.
         * 
         * @param val double value of fraction
         * @param exactDenom the exact denominator
         * @return a SimpleFraction with the given values set.
         */
        public static SimpleFraction BuildFractionExactDenominator(double val, int exactDenom)
        {
            int num = (int)Math.Round(val * (double)exactDenom, MidpointRounding.AwayFromZero);
            return new SimpleFraction(num, exactDenom);
        }

        /**
         * Create a fraction given the double value and either the maximum error
         * allowed or the maximum number of denominator digits.
         *
         * @param value the double value to convert to a fraction.
         * @param maxDenominator maximum denominator value allowed.
         * 
         * @throws RuntimeException if the continued fraction failed to
         *      converge.
         * @throws IllegalArgumentException if value > Integer.MAX_VALUE
         */
        public static SimpleFraction BuildFractionMaxDenominator(double value, int maxDenominator)
        {
            return BuildFractionMaxDenominator(value, 0, maxDenominator, 100);
        }
        /**
         * Create a fraction given the double value and either the maximum error
         * allowed or the maximum number of denominator digits.
         * <p>
         * References:
         * <ul>
         * <li><a href="http://mathworld.wolfram.com/ContinuedFraction.html">
         * Continued Fraction</a> equations (11) and (22)-(26)</li>
         * </ul>
         * </p>
         *
         *  Based on org.apache.commons.math.fraction.Fraction from Apache Commons-Math.
         *  YK: The only reason of having this class is to avoid dependency on the Commons-Math jar.
         *
         * @param value the double value to convert to a fraction.
         * @param epsilon maximum error allowed.  The resulting fraction is within
         *        <code>epsilon</code> of <code>value</code>, in absolute terms.
         * @param maxDenominator maximum denominator value allowed.
         * @param maxIterations maximum number of convergents
         * @throws RuntimeException if the continued fraction failed to
         *         converge.
         * @throws IllegalArgumentException if value > Integer.MAX_VALUE
         */
        private static SimpleFraction BuildFractionMaxDenominator(double value, double epsilon, int maxDenominator, int maxIterations)
        {
            long overflow = long.MaxValue;
            double r0 = value;
            long a0 = (long)Math.Floor(r0);
            if (a0 > overflow)
            {
                throw new ArgumentException("Overflow trying to convert " + value + " to fraction (" + a0 + "/" + 1L + ")");
            }

            // check for (almost) integer arguments, which should not go
            // to iterations.
            if (Math.Abs(a0 - value) < epsilon)
            {
                return new SimpleFraction((int)a0, 1);
            }

            long p0 = 1;
            long q0 = 0;
            long p1 = a0;
            long q1 = 1;

            long p2;
            long q2;

            int n = 0;
            bool stop = false;
            do
            {
                ++n;
                double r1 = 1.0 / (r0 - a0);
                long a1 = (long)Math.Floor(r1);
                p2 = (a1 * p1) + p0;
                q2 = (a1 * q1) + q0;
                //MATH-996/POI-55419
                if (epsilon == 0.0f && maxDenominator > 0 && Math.Abs(q2) > maxDenominator &&
                        Math.Abs(q1) < maxDenominator)
                {

                    return new SimpleFraction((int)p1, (int)q1);
                }
                if ((p2 > overflow) || (q2 > overflow))
                {
                    throw new RuntimeException("Overflow trying to convert " + value + " to fraction (" + p2 + "/" + q2 + ")");
                }

                double convergent = (double)p2 / (double)q2;
                if (n < maxIterations && Math.Abs(convergent - value) > epsilon && q2 < maxDenominator)
                {
                    p0 = p1;
                    p1 = p2;
                    q0 = q1;
                    q1 = q2;
                    a0 = a1;
                    r0 = r1;
                }
                else
                {
                    stop = true;
                }
            } while (!stop);

            if (n >= maxIterations)
            {
                throw new RuntimeException("Unable to convert " + value + " to fraction after " + maxIterations + " iterations");
            }

            if (q2 < maxDenominator)
            {
                return new SimpleFraction((int)p2, (int)q2);
            }
            else
            {
                return new SimpleFraction((int)p1, (int)q1);
            }

        }

        /**
         * Create a fraction given a numerator and denominator.
         * @param numerator
         * @param denominator maxDenominator The maximum allowed value for denominator
         */
        public SimpleFraction(int numerator, int denominator)
        {
            this.numerator = numerator;
            this.denominator = denominator;
        }

        /**
         * Access the denominator.
         * @return the denominator.
         */
        public int Denominator
        {
            get
            {
                return denominator;
            }
        }

        /**
         * Access the numerator.
         * @return the numerator.
         */
        public int Numerator
        {
            get
            {
                return numerator;
            }
        }

    }
}