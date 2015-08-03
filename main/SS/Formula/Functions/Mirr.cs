/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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


namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;

    /**
     * Calculates Modified internal rate of return. Syntax is MIRR(cash_flow_values, finance_rate, reinvest_rate)
     *
     * <p>Returns the modified internal rate of return for a series of periodic cash flows. MIRR considers both the cost
     * of the investment and the interest received on reinvestment of cash.</p>
     *
     * Values is an array or a reference to cells that contain numbers. These numbers represent a series of payments (negative values) and income (positive values) occurring at regular periods.
     * <ul>
     *     <li>Values must contain at least one positive value and one negative value to calculate the modified internal rate of return. Otherwise, MIRR returns the #DIV/0! error value.</li>
     *     <li>If an array or reference argument Contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.</li>
     * </ul>
     *
     * Finance_rate     is the interest rate you pay on the money used in the cash flows.
     * Reinvest_rate     is the interest rate you receive on the cash flows as you reinvest them.
     *
     * @author Carlos Delgado (carlos dot del dot est at gmail dot com)
     * @author Cédric Walter (cedric dot walter at gmail dot com)
     *
     * @see <a href="http://en.wikipedia.org/wiki/MIRR">Wikipedia on MIRR</a>
     * @see <a href="http://office.microsoft.com/en-001/excel-help/mirr-HP005209180.aspx">Excel MIRR</a>
     * @see {@link Irr}
     */
    public class Mirr : MultiOperandNumericFunction
    {

        public Mirr()
            : base(false, false)
        {
        }


        protected override int MaxNumOperands
        {
            get
            {
                return 3;
            }
        }


        protected internal override double Evaluate(double[] values)
        {

            double financeRate = values[values.Length - 1];
            double reinvestRate = values[values.Length - 2];

            double[] mirrValues = new double[values.Length - 2];
            Array.Copy(values, 0, mirrValues, 0, mirrValues.Length);

            bool mirrValuesAreAllNegatives = true;
            foreach (double mirrValue in mirrValues)
            {
                mirrValuesAreAllNegatives &= mirrValue < 0;
            }
            if (mirrValuesAreAllNegatives)
            {
                return -1.0d;
            }

            bool mirrValuesAreAllPositives = true;
            foreach (double mirrValue in mirrValues)
            {
                mirrValuesAreAllPositives &= mirrValue > 0;
            }
            if (mirrValuesAreAllPositives)
            {
                throw new EvaluationException(ErrorEval.DIV_ZERO);
            }

            return mirr(mirrValues, financeRate, reinvestRate);
        }

        private static double mirr(double[] in1, double financeRate, double reinvestRate)
        {
            double value = 0;
            int numOfYears = in1.Length - 1;
            double pv = 0;
            double fv = 0;

            int indexN = 0;
            foreach (double anIn in in1)
            {
                if (anIn < 0)
                {
                    pv += anIn / Math.Pow(1 + financeRate + reinvestRate, indexN++);
                }
            }

            foreach (double anIn in in1)
            {
                if (anIn > 0)
                {
                    fv += anIn * Math.Pow(1 + financeRate, numOfYears - indexN++);
                }
            }

            if (fv != 0 && pv != 0)
            {
                value = Math.Pow(-fv / pv, 1d / numOfYears) - 1;
            }
            return value;
        }
    }

}