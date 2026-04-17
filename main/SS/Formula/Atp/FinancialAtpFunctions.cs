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
using EFF = Excel.FinancialFunctions;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;

namespace NPOI.SS.Formula.Atp
{
    /// <summary>EFFECT(nominal_rate, npery) - effective annual interest rate.</summary>
    internal sealed class EffectFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new EffectFunction();
        private EffectFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 2) return ErrorEval.VALUE_INVALID;
            try
            {
                double nominalRate = FinancialHelper.GetDoubleArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                double npery = FinancialHelper.GetDoubleArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double result = EFF.Financial.Effect(nominalRate, npery);
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>NOMINAL(effect_rate, npery) - nominal annual interest rate.</summary>
    internal sealed class NominalFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new NominalFunction();
        private NominalFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 2) return ErrorEval.VALUE_INVALID;
            try
            {
                double effectRate = FinancialHelper.GetDoubleArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                double npery = FinancialHelper.GetDoubleArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double result = EFF.Financial.Nominal(effectRate, npery);
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>CUMIPMT(rate, nper, pv, start_period, end_period, type)</summary>
    internal sealed class CumIpmtFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new CumIpmtFunction();
        private CumIpmtFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 6) return ErrorEval.VALUE_INVALID;
            try
            {
                double rate = FinancialHelper.GetDoubleArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                double nper = FinancialHelper.GetDoubleArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double pv = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double startPeriod = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                double endPeriod = FinancialHelper.GetDoubleArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                double type = FinancialHelper.GetDoubleArg(args, 5, ec.RowIndex, ec.ColumnIndex);
                double result = EFF.Financial.CumIPmt(rate, nper, pv, startPeriod, endPeriod, FinancialHelper.ToPaymentDue(type));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>CUMPRINC(rate, nper, pv, start_period, end_period, type)</summary>
    internal sealed class CumPrincFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new CumPrincFunction();
        private CumPrincFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 6) return ErrorEval.VALUE_INVALID;
            try
            {
                double rate = FinancialHelper.GetDoubleArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                double nper = FinancialHelper.GetDoubleArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double pv = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double startPeriod = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                double endPeriod = FinancialHelper.GetDoubleArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                double type = FinancialHelper.GetDoubleArg(args, 5, ec.RowIndex, ec.ColumnIndex);
                double result = EFF.Financial.CumPrinc(rate, nper, pv, startPeriod, endPeriod, FinancialHelper.ToPaymentDue(type));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>FVSCHEDULE(principal, schedule)</summary>
    internal sealed class FvScheduleFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new FvScheduleFunction();
        private FvScheduleFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 2) return ErrorEval.VALUE_INVALID;
            try
            {
                double principal = FinancialHelper.GetDoubleArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                double[] schedule = FinancialHelper.GetArrayArg(args, 1);
                double result = EFF.Financial.FvSchedule(principal, schedule);
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>DISC(settlement, maturity, pr, redemption, basis)</summary>
    internal sealed class DiscFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new DiscFunction();
        private DiscFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 4 || args.Length > 5) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double pr = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double redemption = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 4, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.Disc(settlement, maturity, pr, redemption, FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>INTRATE(settlement, maturity, investment, redemption, basis)</summary>
    internal sealed class IntRateFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new IntRateFunction();
        private IntRateFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 4 || args.Length > 5) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double investment = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double redemption = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 4, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.IntRate(settlement, maturity, investment, redemption, FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>RECEIVED(settlement, maturity, investment, discount, basis)</summary>
    internal sealed class ReceivedFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new ReceivedFunction();
        private ReceivedFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 4 || args.Length > 5) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double investment = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double discount = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 4, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.Received(settlement, maturity, investment, discount, FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>DURATION(settlement, maturity, coupon, yld, frequency, basis)</summary>
    internal sealed class DurationFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new DurationFunction();
        private DurationFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 5 || args.Length > 6) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double coupon = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double yld = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 5, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.Duration(settlement, maturity, coupon, yld, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>MDURATION(settlement, maturity, coupon, yld, frequency, basis)</summary>
    internal sealed class MDurationFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new MDurationFunction();
        private MDurationFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 5 || args.Length > 6) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double coupon = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double yld = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 5, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.MDuration(settlement, maturity, coupon, yld, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>PRICE(settlement, maturity, rate, yld, redemption, frequency, basis)</summary>
    internal sealed class PriceFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new PriceFunction();
        private PriceFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 6 || args.Length > 7) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double rate = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double yld = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                double redemption = FinancialHelper.GetDoubleArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 5, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 6, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.Price(settlement, maturity, rate, yld, redemption, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>PRICEDISC(settlement, maturity, discount, redemption, basis)</summary>
    internal sealed class PriceDiscFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new PriceDiscFunction();
        private PriceDiscFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 4 || args.Length > 5) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double discount = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double redemption = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 4, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.PriceDisc(settlement, maturity, discount, redemption, FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>PRICEMAT(settlement, maturity, issue, rate, yld, basis)</summary>
    internal sealed class PriceMatFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new PriceMatFunction();
        private PriceMatFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 5 || args.Length > 6) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                DateTime issue = FinancialHelper.GetDateArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double rate = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                double yld = FinancialHelper.GetDoubleArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 5, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.PriceMat(settlement, maturity, issue, rate, yld, FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>YIELD(settlement, maturity, rate, pr, redemption, frequency, basis)</summary>
    internal sealed class YieldFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new YieldFunction();
        private YieldFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 6 || args.Length > 7) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double rate = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double pr = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                double redemption = FinancialHelper.GetDoubleArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 5, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 6, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.Yield(settlement, maturity, rate, pr, redemption, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>YIELDDISC(settlement, maturity, pr, redemption, basis)</summary>
    internal sealed class YieldDiscFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new YieldDiscFunction();
        private YieldDiscFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 4 || args.Length > 5) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double pr = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double redemption = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 4, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.YieldDisc(settlement, maturity, pr, redemption, FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>YIELDMAT(settlement, maturity, issue, rate, pr, basis)</summary>
    internal sealed class YieldMatFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new YieldMatFunction();
        private YieldMatFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 5 || args.Length > 6) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                DateTime issue = FinancialHelper.GetDateArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double rate = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                double pr = FinancialHelper.GetDoubleArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 5, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.YieldMat(settlement, maturity, issue, rate, pr, FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>TBILLEQ(settlement, maturity, discount)</summary>
    internal sealed class TBillEqFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new TBillEqFunction();
        private TBillEqFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 3) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double discount = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double result = EFF.Financial.TBillEq(settlement, maturity, discount);
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>TBILLPRICE(settlement, maturity, discount)</summary>
    internal sealed class TBillPriceFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new TBillPriceFunction();
        private TBillPriceFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 3) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double discount = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double result = EFF.Financial.TBillPrice(settlement, maturity, discount);
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>TBILLYIELD(settlement, maturity, pr)</summary>
    internal sealed class TBillYieldFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new TBillYieldFunction();
        private TBillYieldFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 3) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double pr = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double result = EFF.Financial.TBillYield(settlement, maturity, pr);
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>COUPDAYBS(settlement, maturity, frequency, basis)</summary>
    internal sealed class CoupDayBsFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new CoupDayBsFunction();
        private CoupDayBsFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 3 || args.Length > 4) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 3, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.CoupDaysBS(settlement, maturity, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>COUPDAYS(settlement, maturity, frequency, basis)</summary>
    internal sealed class CoupDaysFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new CoupDaysFunction();
        private CoupDaysFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 3 || args.Length > 4) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 3, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.CoupDays(settlement, maturity, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>COUPDAYSNC(settlement, maturity, frequency, basis)</summary>
    internal sealed class CoupDaysNcFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new CoupDaysNcFunction();
        private CoupDaysNcFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 3 || args.Length > 4) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 3, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.CoupDaysNC(settlement, maturity, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>COUPNCD(settlement, maturity, frequency, basis) - next coupon date after settlement.</summary>
    internal sealed class CoupNcdFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new CoupNcdFunction();
        private CoupNcdFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 3 || args.Length > 4) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 3, ec.RowIndex, ec.ColumnIndex, 0);
                DateTime result = EFF.Financial.CoupNCD(settlement, maturity, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis));
                // Convert DateTime back to Excel serial date number.
                // Compensate for the Lotus 1-2-3 leap-year bug: Excel treats 1900 as a
                // leap year, inserting a phantom Feb 29 (serial 60). Dates on or after
                // Mar 1 1900 are serialised one day higher than the actual count from epoch.
                int serial = (result - new DateTime(1899, 12, 31)).Days;
                if (result >= new DateTime(1900, 3, 1))
                    serial += 1;
                NumericFunction.CheckValue(serial);
                return new NumberEval(serial);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>COUPNUM(settlement, maturity, frequency, basis)</summary>
    internal sealed class CoupNumFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new CoupNumFunction();
        private CoupNumFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 3 || args.Length > 4) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 3, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.CoupNum(settlement, maturity, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>COUPPCD(settlement, maturity, frequency, basis) - previous coupon date before settlement.</summary>
    internal sealed class CoupPcdFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new CoupPcdFunction();
        private CoupPcdFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 3 || args.Length > 4) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 3, ec.RowIndex, ec.ColumnIndex, 0);
                DateTime result = EFF.Financial.CoupPCD(settlement, maturity, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis));
                int serial = (result - new DateTime(1899, 12, 31)).Days;
                if (result >= new DateTime(1900, 3, 1))
                    serial += 1;
                NumericFunction.CheckValue(serial);
                return new NumberEval(serial);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>ACCRINT(issue, first_interest, settlement, rate, par, frequency, basis, calc_method)</summary>
    internal sealed class AccrintFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new AccrintFunction();
        private AccrintFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 6 || args.Length > 8) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime issue = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime firstInterest = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                DateTime settlement = FinancialHelper.GetDateArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double rate = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                double par = FinancialHelper.GetDoubleArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 5, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 6, ec.RowIndex, ec.ColumnIndex, 0);
                double calcMethod = FinancialHelper.GetDoubleArg(args, 7, ec.RowIndex, ec.ColumnIndex, 1.0);
                EFF.AccrIntCalcMethod method = calcMethod != 0.0 ? EFF.AccrIntCalcMethod.FromIssueToSettlement : EFF.AccrIntCalcMethod.FromFirstToSettlement;
                double result = EFF.Financial.AccrInt(issue, firstInterest, settlement, rate, par, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis), method);
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>ACCRINTM(issue, settlement, rate, par, basis)</summary>
    internal sealed class AccrintMFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new AccrintMFunction();
        private AccrintMFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 4 || args.Length > 5) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime issue = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime settlement = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                double rate = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double par = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 4, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.AccrIntM(issue, settlement, rate, par, FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>AMORDEGRC(cost, date_purchased, first_period, salvage, period, rate, basis)</summary>
    internal sealed class AmorDegrcFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new AmorDegrcFunction();
        private AmorDegrcFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 6 || args.Length > 7) return ErrorEval.VALUE_INVALID;
            try
            {
                double cost = FinancialHelper.GetDoubleArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime datePurchased = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                DateTime firstPeriod = FinancialHelper.GetDateArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double salvage = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                double period = FinancialHelper.GetDoubleArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                double rate = FinancialHelper.GetDoubleArg(args, 5, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 6, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.AmorDegrc(cost, datePurchased, firstPeriod, salvage, period, rate, FinancialHelper.ToDayCountBasis(basis), true);
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>AMORLINC(cost, date_purchased, first_period, salvage, period, rate, basis)</summary>
    internal sealed class AmorLincFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new AmorLincFunction();
        private AmorLincFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 6 || args.Length > 7) return ErrorEval.VALUE_INVALID;
            try
            {
                double cost = FinancialHelper.GetDoubleArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime datePurchased = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                DateTime firstPeriod = FinancialHelper.GetDateArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double salvage = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                double period = FinancialHelper.GetDoubleArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                double rate = FinancialHelper.GetDoubleArg(args, 5, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 6, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.AmorLinc(cost, datePurchased, firstPeriod, salvage, period, rate, FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld, redemption, frequency, basis)</summary>
    internal sealed class OddFPriceFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new OddFPriceFunction();
        private OddFPriceFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 8 || args.Length > 9) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                DateTime issue = FinancialHelper.GetDateArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                DateTime firstCoupon = FinancialHelper.GetDateArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                double rate = FinancialHelper.GetDoubleArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                double yld = FinancialHelper.GetDoubleArg(args, 5, ec.RowIndex, ec.ColumnIndex);
                double redemption = FinancialHelper.GetDoubleArg(args, 6, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 7, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 8, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.OddFPrice(settlement, maturity, issue, firstCoupon, rate, yld, redemption, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr, redemption, frequency, basis)</summary>
    internal sealed class OddFYieldFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new OddFYieldFunction();
        private OddFYieldFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 8 || args.Length > 9) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                DateTime issue = FinancialHelper.GetDateArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                DateTime firstCoupon = FinancialHelper.GetDateArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                double rate = FinancialHelper.GetDoubleArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                double pr = FinancialHelper.GetDoubleArg(args, 5, ec.RowIndex, ec.ColumnIndex);
                double redemption = FinancialHelper.GetDoubleArg(args, 6, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 7, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 8, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.OddFYield(settlement, maturity, issue, firstCoupon, rate, pr, redemption, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>ODDLPRICE(settlement, maturity, last_interest, rate, yld, redemption, frequency, basis)</summary>
    internal sealed class OddLPriceFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new OddLPriceFunction();
        private OddLPriceFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 7 || args.Length > 8) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                DateTime lastInterest = FinancialHelper.GetDateArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double rate = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                double yld = FinancialHelper.GetDoubleArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                double redemption = FinancialHelper.GetDoubleArg(args, 5, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 6, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 7, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.OddLPrice(settlement, maturity, lastInterest, rate, yld, redemption, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>ODDLYIELD(settlement, maturity, last_interest, rate, pr, redemption, frequency, basis)</summary>
    internal sealed class OddLYieldFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new OddLYieldFunction();
        private OddLYieldFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 7 || args.Length > 8) return ErrorEval.VALUE_INVALID;
            try
            {
                DateTime settlement = FinancialHelper.GetDateArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                DateTime maturity = FinancialHelper.GetDateArg(args, 1, ec.RowIndex, ec.ColumnIndex);
                DateTime lastInterest = FinancialHelper.GetDateArg(args, 2, ec.RowIndex, ec.ColumnIndex);
                double rate = FinancialHelper.GetDoubleArg(args, 3, ec.RowIndex, ec.ColumnIndex);
                double pr = FinancialHelper.GetDoubleArg(args, 4, ec.RowIndex, ec.ColumnIndex);
                double redemption = FinancialHelper.GetDoubleArg(args, 5, ec.RowIndex, ec.ColumnIndex);
                int frequency = FinancialHelper.GetIntArg(args, 6, ec.RowIndex, ec.ColumnIndex);
                int basis = FinancialHelper.GetIntArg(args, 7, ec.RowIndex, ec.ColumnIndex, 0);
                double result = EFF.Financial.OddLYield(settlement, maturity, lastInterest, rate, pr, redemption, FinancialHelper.ToFrequency(frequency), FinancialHelper.ToDayCountBasis(basis));
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>XIRR(values, dates[, guess])</summary>
    internal sealed class XirrFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new XirrFunction();
        private XirrFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 2 || args.Length > 3) return ErrorEval.VALUE_INVALID;
            try
            {
                double[] values = FinancialHelper.GetArrayArg(args, 0);
                double[] dateSerials = FinancialHelper.GetArrayArg(args, 1);
                double guess = FinancialHelper.GetDoubleArg(args, 2, ec.RowIndex, ec.ColumnIndex, 0.1);

                if (values.Length != dateSerials.Length)
                    return ErrorEval.VALUE_INVALID;

                DateTime[] dates = new DateTime[dateSerials.Length];
                for (int i = 0; i < dateSerials.Length; i++)
                    dates[i] = FinancialHelper.FromExcelDate(dateSerials[i]);

                double result = EFF.Financial.XIrr(values, dates, guess);
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }

    /// <summary>XNPV(rate, values, dates)</summary>
    internal sealed class XnpvFunction : FreeRefFunction
    {
        public static readonly FreeRefFunction instance = new XnpvFunction();
        private XnpvFunction() { }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 3) return ErrorEval.VALUE_INVALID;
            try
            {
                double rate = FinancialHelper.GetDoubleArg(args, 0, ec.RowIndex, ec.ColumnIndex);
                double[] values = FinancialHelper.GetArrayArg(args, 1);
                double[] dateSerials = FinancialHelper.GetArrayArg(args, 2);

                if (values.Length != dateSerials.Length)
                    return ErrorEval.VALUE_INVALID;

                DateTime[] dates = new DateTime[dateSerials.Length];
                for (int i = 0; i < dateSerials.Length; i++)
                    dates[i] = FinancialHelper.FromExcelDate(dateSerials[i]);

                double result = EFF.Financial.XNpv(rate, values, dates);
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e) { return e.GetErrorEval(); }
        }
    }
}
