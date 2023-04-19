using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Atp
{
    class XLookupFunction : FreeRefFunction, ArrayFunction
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
            return _evaluate(args, srcRowIndex, srcColumnIndex, ec.IsSingleValue);
        }

        public ValueEval EvaluateArray(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            return _evaluate(args, srcRowIndex, srcColumnIndex, false);
        }

        private String LaxValueToString(ValueEval eval)
        {
            return (eval is MissingArgEval) ? "" : OperandResolver.CoerceValueToString(eval);
        }
        private ValueEval _evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex, bool isSingleValue)
        {
            if (args.Length < 3)
            {
                return ErrorEval.VALUE_INVALID;
            }
            String notFound = null;
            if (args.Length > 3)
            {
                try
                {
                    ValueEval notFoundValue = OperandResolver.GetSingleValue(args[3], srcRowIndex, srcColumnIndex);
                    String notFoundText = LaxValueToString(notFoundValue);
                    if (notFoundText != null)
                    {
                        String trimmedText = notFoundText.Trim();
                        if (trimmedText.Length > 0)
                        {
                            notFound = trimmedText;
                        }
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
            return evaluate(srcRowIndex, srcColumnIndex, args[0], args[1], args[2], notFound, matchMode, searchMode, isSingleValue);
        }
        private ValueEval evaluate(int srcRowIndex, int srcColumnIndex, ValueEval lookupEval, ValueEval indexEval,
                           ValueEval returnEval, String notFound, LookupUtils.MatchMode matchMode,
                           LookupUtils.SearchMode searchMode, bool isSingleValue)
        {
            try
            {
                ValueEval lookupValue = OperandResolver.GetSingleValue(lookupEval, srcRowIndex, srcColumnIndex);
                TwoDEval tableArray = LookupUtils.ResolveTableArrayArg(indexEval);
                int matchedRow;
                try
                {
                    matchedRow = LookupUtils.XlookupIndexOfValue(lookupValue, LookupUtils.CreateColumnVector(tableArray, 0), matchMode, searchMode);
                }
                catch (EvaluationException e)
                {
                    if (ErrorEval.NA.Equals(e.GetErrorEval()))
                    {
                        if (string.IsNullOrEmpty(notFound))
                        {
                            if (returnEval is AreaEval) {
                                AreaEval area = (AreaEval)returnEval;
                                int width = area.Width;
                                if (isSingleValue || width <= 1)
                                {
                                    return new StringEval(notFound);
                                }
                                return new NotFoundAreaEval(notFound, width);
                            } else
                            {
                                return new StringEval(notFound);
                            }
                        }
                        return ErrorEval.NA;
                    }
                    else
                    {
                        return e.GetErrorEval();
                    }
                }
                if (returnEval is AreaEval) {
                    AreaEval area = (AreaEval)returnEval;
                    if (isSingleValue)
                    {
                        return area.GetRelativeValue(matchedRow, 0);
                    }
                    return area.Offset(matchedRow, matchedRow, 0, area.Width - 1);
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
            private int _width;
            private string _notFound;
            public NotFoundAreaEval(string notFound, int width)
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
                    return new StringEval(_notFound);
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

            public AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx)
            {
                throw null;
            }
        }
    }
}
