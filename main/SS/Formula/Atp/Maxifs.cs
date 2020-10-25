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


namespace NPOI.SS.Formula.Atp
{
    using System;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula.Eval;

    /**
     * Implementation for the function MAXIFS
     * <p>
     * Syntax: MAXIFS(data_range, criteria_range1, criteria1, [criteria_range2, criteria2])
     * </p>
     */

    public class Maxifs : FreeRefFunction
    {
        public static FreeRefFunction instance = new Maxifs();

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 3 || args.Length % 2 == 0)
            {
                return ErrorEval.VALUE_INVALID;
            }

            try
            {
                AreaEval dataRange = Sumifs.ConvertRangeArg(args[0]);

                // collect pairs of ranges and criteria
                AreaEval[] ae = new AreaEval[(args.Length - 1) / 2];
                IMatchPredicate[] mp = new IMatchPredicate[ae.Length];
                for (int i = 1, k = 0; i < args.Length; i += 2, k++)
                {
                    ae[k] = Sumifs.ConvertRangeArg(args[i]);
                    mp[k] = Countif.CreateCriteriaPredicate(args[i + 1], ec.RowIndex, ec.ColumnIndex);
                }

                Sumifs.ValidateCriteriaRanges(ae, dataRange);
                Sumifs.ValidateCriteria(mp);

                double result = Sumifs.CalcMatchingCells(ae, mp, dataRange, double.NaN, (init, current) => !current.HasValue ? init : double.IsNaN(init) ? current.Value : current.Value > init ? current.Value : init);
                return new NumberEval(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
    }

}
