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
 * Created on May 21, 2005
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using System;
    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     * 
     * 
     * This class Is a functon library for common fiscal functions.
     * <b>Glossary of terms/abbreviations:</b>
     * <br/>
     * <ul>
     * <li><em>FV:</em> Future Value</li>
     * <li><em>PV:</em> Present Value</li>
     * <li><em>NPV:</em> Net Present Value</li>
     * <li><em>PMT:</em> (Periodic) Payment</li>
     * 
     * </ul>
     * For more info on the terms/abbreviations please use the references below 
     * (hyperlinks are subject to Change):
     * <br/>Online References:
     * <ol>
     * <li>GNU Emacs Calc 2.02 Manual: http://theory.uwinnipeg.ca/gnu/calc/calc_203.html</li>
     * <li>Yahoo Financial Glossary: http://biz.yahoo.com/f/g/nn.html#y</li>
     * <li>MS Excel function reference: http://office.microsoft.com/en-us/assistance/CH062528251033.aspx</li>
     * </ol>
     * <h3>Implementation Notes:</h3>
     * Symbols used in the formulae that follow:<br/>
     * <ul>
     * <li>p: present value</li>
     * <li>f: future value</li>
     * <li>n: number of periods</li>
     * <li>y: payment (in each period)</li>
     * <li>r: rate</li>
     * <li>^: the power operator (NOT the java bitwise XOR operator!)</li>
     * </ul>
     * [From MS Excel function reference] Following are some of the key formulas
     * that are used in this implementation:
     * <pre>
     * p(1+r)^n + y(1+rt)((1+r)^n-1)/r + f=0   ...{when r!=0}
     * ny + p + f=0                            ...{when r=0}
     * </pre>
     */
    public class FinanceLib
    {
        
        // constants for default values



        private FinanceLib() { }

        /**
         * Future value of an amount given the number of payments, rate, amount
         * of individual payment, present value and bool value indicating whether
         * payments are due at the beginning of period 
         * (false => payments are due at end of period) 
         * @param r rate
         * @param n num of periods
         * @param y pmt per period
         * @param p future value
         * @param t type (true=pmt at end of period, false=pmt at begining of period)
         */
        public static double fv(double r, double n, double y, double p, bool t)
        {
            double retval = 0;
            if (r == 0)
            {
                retval = -1 * (p + (n * y));
            }
            else
            {
                double r1 = r + 1;
                retval = ((1 - Math.Pow(r1, n)) * (t ? r1 : 1) * y) / r
                          -
                       p * Math.Pow(r1, n);
            }
            return retval;
        }

        /**
         * Present value of an amount given the number of future payments, rate, amount
         * of individual payment, future value and bool value indicating whether
         * payments are due at the beginning of period 
         * (false => payments are due at end of period) 
         * @param r
         * @param n
         * @param y
         * @param f
         * @param t
         */
        public static double pv(double r, double n, double y, double f, bool t)
        {
            double retval = 0;
            if (r == 0)
            {
                retval = -1 * ((n * y) + f);
            }
            else
            {
                double r1 = r + 1;
                retval = (((1 - Math.Pow(r1, n)) / r) * (t ? r1 : 1) * y - f)
                         /
                        Math.Pow(r1, n);
            }
            return retval;
        }

        /**
         * calculates the Net Present Value of a principal amount
         * given the disCount rate and a sequence of cash flows 
         * (supplied as an array). If the amounts are income the value should 
         * be positive, else if they are payments and not income, the 
         * value should be negative.
         * @param r
         * @param cfs cashflow amounts
         */
        public static double npv(double r, double[] cfs)
        {
            double npv = 0;
            double r1 = r + 1;
            double trate = r1;
            for (int i = 0, iSize = cfs.Length; i < iSize; i++)
            {
                npv += cfs[i] / trate;
                trate *= r1;
            }
            return npv;
        }

        /**
         * 
         * @param r
         * @param n
         * @param p
         * @param f
         * @param t
         */
        public static double pmt(double r, double n, double p, double f, bool t)
        {
            double retval = 0;
            if (r == 0)
            {
                retval = -1 * (f + p) / n;
            }
            else
            {
                double r1 = r + 1;
                retval = (f + p * Math.Pow(r1, n)) * r
                          /
                       ((t ? r1 : 1) * (1 - Math.Pow(r1, n)));
            }
            return retval;
        }

        /**
         * 
         * @param r
         * @param y
         * @param p
         * @param f
         * @param t
         */
        public static double nper(double r, double y, double p, double f, bool t)
        {
            double retval = 0;
            if (r == 0)
            {
                retval = -1 * (f + p) / y;
            }
            else
            {
                double r1 = r + 1;
                double ryr = (t ? r1 : 1) * y / r;
                double a1 = ((ryr - f) < 0)
                        ? Math.Log(f - ryr)
                        : Math.Log(ryr - f);
                double a2 = ((ryr - f) < 0)
                        ? Math.Log(-p - ryr)
                        : Math.Log(p + ryr);
                double a3 = Math.Log(r1);
                retval = (a1 - a2) / a3;
            }
            return retval;
        }


    }
}