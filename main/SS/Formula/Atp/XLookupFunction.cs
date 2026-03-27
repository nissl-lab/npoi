using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Atp
{
    class XLookupFunction : FreeRefFunction, IArrayFunction
    {
        public static FreeRefFunction instance = new XLookupFunction(ArgumentsEvaluator.instance);

        private ArgumentsEvaluator evaluator;

        private XLookupFunction(ArgumentsEvaluator anEvaluator)
        {
            // enforces singleton
            this.evaluator = anEvaluator;
        }
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            int srcRowIndex = ec.RowIndex;
            int srcColumnIndex = ec.ColumnIndex;
            return _evaluate(args, srcRowIndex, srcColumnIndex);
        }

        public ValueEval EvaluateArray(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            return _evaluate(args, srcRowIndex, srcColumnIndex);
        }

        private static String LaxValueToString(ValueEval eval)
        {
            return (eval is MissingArgEval) ? "" : OperandResolver.CoerceValueToString(eval);
        }

        private static ValueEval _evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length < 3)
            {
                return ErrorEval.VALUE_INVALID;
            }
            ValueEval notFound = BlankEval.instance;
            if (args.Length > 3)
            {
                try
                {
                    ValueEval notFoundValue = OperandResolver.GetSingleValue(args[3], srcRowIndex, srcColumnIndex);
                    if (notFoundValue != null)
                    {
                        notFound = notFoundValue;
                    }
                }
                catch (EvaluationException e)
                {
                    return e.GetErrorEval();
                }
            }
            LookupUtils.MatchMode matchMode = LookupUtils.MatchMode.ExactMatch;
            if (args.Length > 4)
            {
                try
                {
                    ValueEval matchModeValue = OperandResolver.GetSingleValue(args[4], srcRowIndex, srcColumnIndex);
                    int matchInt = OperandResolver.CoerceValueToInt(matchModeValue);
                    matchMode = LookupUtils.GetMatchMode(matchInt);
                }
                catch (EvaluationException e)
                {
                    return e.GetErrorEval();
                }
                catch (Exception )
                {
                    return ErrorEval.VALUE_INVALID;
                }
            }
            LookupUtils.SearchMode searchMode = LookupUtils.SearchMode.IterateForward;
            if (args.Length > 5)
            {
                try
                {
                    ValueEval searchModeValue = OperandResolver.GetSingleValue(args[5], srcRowIndex, srcColumnIndex);
                    int searchInt = OperandResolver.CoerceValueToInt(searchModeValue);
                    searchMode = LookupUtils.GetSearchMode(searchInt);
                }
                catch (EvaluationException e)
                {
                    return e.GetErrorEval();
                }
                catch
                {
                    return ErrorEval.VALUE_INVALID;
                }
            }
            return evaluate(srcRowIndex, srcColumnIndex, args[0], args[1], args[2], notFound, matchMode, searchMode);
        }

        private static ValueEval evaluate(int srcRowIndex, int srcColumnIndex, ValueEval lookupEval, ValueEval indexEval,
                           ValueEval returnEval, ValueEval notFound, LookupUtils.MatchMode matchMode,
                           LookupUtils.SearchMode searchMode)
        {
            try
            {
                ValueEval lookupValue = OperandResolver.GetSingleValue(lookupEval, srcRowIndex, srcColumnIndex);
                TwoDEval tableArray = LookupUtils.ResolveTableArrayArg(indexEval);
                ValueVector vector;
                if (tableArray.IsColumn) {
                    vector = LookupUtils.CreateColumnVector(tableArray, 0);
                } else {
                    vector = LookupUtils.CreateColumnVector(tableArray, 0);
                }
                int matchedIdx;
                try
                {
                    matchedIdx = LookupUtils.XlookupIndexOfValue(lookupValue, vector, matchMode, searchMode);
                }
                catch (EvaluationException e)
                {
                    if (ErrorEval.NA.Equals(e.GetErrorEval()))
                    {
                        if (notFound != BlankEval.instance)
                        {
                            if (returnEval is AreaEval area) {
                                int width = area.Width;
                                if (width <= 1)
                                {
                                    return notFound;
                                }
                                return new NotFoundAreaEval(notFound, width);
                            } else
                            {
                                return notFound;
                            }
                        }
                        return ErrorEval.NA;
                    }
                    else
                    {
                        return e.GetErrorEval();
                    }
                }
                if (returnEval is AreaEval eval) {
                    AreaEval area = (AreaEval)returnEval;
                    if (tableArray.IsColumn) {
                        return area.Offset(matchedIdx, matchedIdx,0, area.Width - 1);
                    } else {
                        return area.Offset(0, area.Height - 1,matchedIdx, matchedIdx);
                    }
                } else
                {
                    return returnEval;
                }
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        class NotFoundAreaEval : AreaEval
        {
            private readonly int _width;
            private readonly ValueEval _notFound;
            public NotFoundAreaEval(ValueEval notFound, int width)
            {
                _width = width;
                _notFound = notFound;
            }
            public int FirstRow 
            { 
                get { return 0; } 
            }

            public int LastRow
            {
                get { return 0; }
            }

            public int FirstColumn
            {
                get { return 0; }
            }

            public int LastColumn
            {
                get { return _width - 1; }
            }

            public int Width
            {
                get { return _width; }
            }

            public int Height
            {
                get { return 1; }
            }

            public bool IsRow
            {
                get { return false; }
            }

            public bool IsColumn
            {
                get { return false; }
            }

            public int FirstSheetIndex
            {
                get { return 0; }
            }

            public int LastSheetIndex
            {
                get { return 0; }
            }

            public bool Contains(int row, int col)
            {
                throw new NotImplementedException();
            }

            public bool ContainsColumn(int col)
            {
                throw new NotImplementedException();
            }

            public bool ContainsRow(int row)
            {
                throw new NotImplementedException();
            }

            public ValueEval GetAbsoluteValue(int row, int col)
            {
                if (col == 0)
                {
                    return _notFound;
                }
                return new StringEval("");
            }

            public TwoDEval GetColumn(int columnIndex)
            {
                throw null;
            }

            public ValueEval GetRelativeValue(int relativeRowIndex, int relativeColumnIndex)
            {
                return GetAbsoluteValue(relativeRowIndex, relativeColumnIndex);
            }

            public TwoDEval GetRow(int rowIndex)
            {
                throw null;
            }

            public ValueEval GetValue(int sheetIndex, int rowIndex, int columnIndex)
            {
                return GetAbsoluteValue(rowIndex, columnIndex);
            }

            public ValueEval GetValue(int rowIndex, int columnIndex)
            {
                return GetAbsoluteValue(rowIndex, columnIndex);
            }

            public bool IsSubTotal(int rowIndex, int columnIndex)
            {
                return false;
            }

            public bool IsRowHidden(int rowIndex)
            {
                return false;
            }
            public AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx)
            {
                throw null;
            }
        }
    }
}
