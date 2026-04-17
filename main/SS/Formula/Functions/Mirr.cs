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
    using EFF = Excel.FinancialFunctions;
    using NPOI.SS.Formula.Eval;

    public class Mirr : MultiOperandNumericFunction
    {
        public Mirr()
            : base(false, false)
        {
        }

        internal override int MaxNumOperands
        {
            get { return 3; }
        }

        protected internal override double Evaluate(double[] values)
        {
            if (values.Length < 3)
                throw new EvaluationException(ErrorEval.VALUE_INVALID);

            double financeRate = values[values.Length - 1];
            double reinvestRate = values[values.Length - 2];

            double[] mirrValues = new double[values.Length - 2];
            Array.Copy(values, 0, mirrValues, 0, mirrValues.Length);

            bool mirrValuesAreAllNegatives = true;
            foreach (double mirrValue in mirrValues)
                mirrValuesAreAllNegatives &= mirrValue < 0;
            if (mirrValuesAreAllNegatives)
                return -1.0d;

            bool mirrValuesAreAllPositives = true;
            foreach (double mirrValue in mirrValues)
                mirrValuesAreAllPositives &= mirrValue > 0;
            if (mirrValuesAreAllPositives)
                throw new EvaluationException(ErrorEval.DIV_ZERO);

            try
            {
                return EFF.Financial.Mirr(mirrValues, reinvestRate, financeRate);
            }
            catch (Exception)
            {
                throw new EvaluationException(ErrorEval.NUM_ERROR);
            }
        }
    }
}