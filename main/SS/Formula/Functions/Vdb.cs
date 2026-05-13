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

using EFF = Excel.FinancialFunctions;
using NPOI.SS.Formula.Eval;

namespace NPOI.SS.Formula.Functions
{
    /// <summary>
    /// Implements the Excel VDB() function.
    /// VDB(cost, salvage, life, start_period, end_period[, factor[, no_switch]]) - variable declining balance depreciation.
    /// </summary>
    public class Vdb : NumericFunction
    {
        protected override double Eval(ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            if (args.Length < 5 || args.Length > 7)
                throw new EvaluationException(ErrorEval.VALUE_INVALID);

            double cost = FinancialHelper.GetDoubleArg(args, 0, srcCellRow, srcCellCol);
            double salvage = FinancialHelper.GetDoubleArg(args, 1, srcCellRow, srcCellCol);
            double life = FinancialHelper.GetDoubleArg(args, 2, srcCellRow, srcCellCol);
            double startPeriod = FinancialHelper.GetDoubleArg(args, 3, srcCellRow, srcCellCol);
            double endPeriod = FinancialHelper.GetDoubleArg(args, 4, srcCellRow, srcCellCol);
            double factor = FinancialHelper.GetDoubleArg(args, 5, srcCellRow, srcCellCol, 2.0);
            double noSwitch = FinancialHelper.GetDoubleArg(args, 6, srcCellRow, srcCellCol, 0.0);

            EFF.VdbSwitch vdbSwitch = noSwitch != 0.0 ? EFF.VdbSwitch.DontSwitchToStraightLine : EFF.VdbSwitch.SwitchToStraightLine;
            double result = EFF.Financial.Vdb(cost, salvage, life, startPeriod, endPeriod, factor, vdbSwitch);
            CheckValue(result);
            return result;
        }
    }
}
