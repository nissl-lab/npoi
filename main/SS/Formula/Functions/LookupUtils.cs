/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.SS.Formula.Functions
{
    using System;
    using System.Text;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using System.Globalization;
    using System.Text.RegularExpressions;
    using System.Collections.Generic;
    using EnumsNET;

    /**
     * Common functionality used by VLOOKUP, HLOOKUP, LOOKUP and MATCH
     * 
     * @author Josh Micich
     */
    internal sealed class LookupUtils
    {
        private static readonly Dictionary<int, MatchMode> matchModeMap = new Dictionary<int, MatchMode>();
        private static readonly Dictionary<int, SearchMode> searchModeMap = new Dictionary<int, SearchMode>();

        static LookupUtils()
        {
            foreach (var value in Enums.GetValues<MatchMode>())
            {
                matchModeMap.Add((int)value, value);
            }
            foreach (var value in Enums.GetValues<SearchMode>())
            {
                searchModeMap.Add((int)value, value);
            }
        }
        public static MatchMode GetMatchMode(int m)
        {
            if (!matchModeMap.TryGetValue(m, out MatchMode mode))
            {
                throw new ArgumentException("unknown match mode " + m);
            }
            return mode;
        }
        public static SearchMode GetSearchMode(int s)
        {
            if (!searchModeMap.TryGetValue(s, out SearchMode mode))
            {
                throw new ArgumentException("unknown search mode " + s);
            }
            return mode;
        }
        public enum MatchMode : int
        {
            ExactMatch = 0,
            ExactMatchFallbackToSmallerValue = -1,
            ExactMatchFallbackToLargerValue = 1,
            WildcardMatch = 2
        }
        public enum SearchMode
        {
            IterateForward = 1,
            IterateBackward = -1,
            BinarySearchForward = 2,
            BinarySearchBackward = -2
        }
        internal sealed class RowVector : ValueVector
        {

            private readonly AreaEval _tableArray;
            private readonly int _size;
            private readonly int _rowIndex;

            public RowVector(AreaEval tableArray, int rowIndex)
            {
                _rowIndex = rowIndex;
                int _rowAbsoluteIndex = tableArray.FirstRow + rowIndex;
                if (!tableArray.ContainsRow(_rowAbsoluteIndex))
                {
                    int lastRowIx = tableArray.LastRow - tableArray.FirstRow;
                    throw new ArgumentException("Specified row index (" + rowIndex
                            + ") is outside the allowed range (0.." + lastRowIx + ")");
                }
                _tableArray = tableArray;
                _size = tableArray.Width;
            }

            public ValueEval GetItem(int index)
            {
                if (index > _size)
                {
                    throw new IndexOutOfRangeException("Specified index (" + index
                            + ") is outside the allowed range (0.." + (_size - 1) + ")");
                }
                return _tableArray.GetRelativeValue(_rowIndex, index);
            }
            public int Size
            {
                get
                {
                    return _size;
                }
            }
        }
        internal sealed class ColumnVector : ValueVector
        {

            private readonly AreaEval _tableArray;
            private readonly int _size;
            private readonly int _columnIndex;

            public ColumnVector(AreaEval tableArray, int columnIndex)
            {
                _columnIndex = columnIndex;
                int _columnAbsoluteIndex = tableArray.FirstColumn + columnIndex;
                if (!tableArray.ContainsColumn((short)_columnAbsoluteIndex))
                {
                    int lastColIx = tableArray.LastColumn - tableArray.FirstColumn;
                    throw new ArgumentException("Specified column index (" + columnIndex
                            + ") is outside the allowed range (0.." + lastColIx + ")");
                }
                _tableArray = tableArray;
                _size = _tableArray.Height;
            }

            public ValueEval GetItem(int index)
            {
                if (index > _size)
                {
                    throw new IndexOutOfRangeException("Specified index (" + index
                            + ") is outside the allowed range (0.." + (_size - 1) + ")");
                }
                return _tableArray.GetRelativeValue(index, _columnIndex);
            }
            public int Size
            {
                get
                {
                    return _size;
                }
            }
        }

        private sealed class SheetVector : ValueVector
        {
            private readonly RefEval _re;
            private readonly int _size;

            public SheetVector(RefEval re)
            {
                _size = re.NumberOfSheets;
                _re = re;
            }

            public ValueEval GetItem(int index)
            {
                if (index >= _size)
                {
                    throw new IndexOutOfRangeException("Specified index (" + index
                            + ") is outside the allowed range (0.." + (_size - 1) + ")");
                }
                int sheetIndex = _re.FirstSheetIndex + index;
                return _re.GetInnerValueEval(sheetIndex);
            }
            public int Size
            {
                get
                {
                    return _size;
                }
            }
        }


        public static ValueVector CreateRowVector(TwoDEval tableArray, int relativeRowIndex)
        {
            return new RowVector((AreaEval)tableArray, relativeRowIndex);
        }
        public static ValueVector CreateColumnVector(TwoDEval tableArray, int relativeColumnIndex)
        {
            return new ColumnVector((AreaEval)tableArray, relativeColumnIndex);
        }
        /**
         * @return <c>null</c> if the supplied area is neither a single row nor a single colum
         */
        public static ValueVector CreateVector(TwoDEval ae)
        {
            if (ae.IsColumn)
            {
                return CreateColumnVector(ae, 0);
            }
            if (ae.IsRow)
            {
                return CreateRowVector(ae, 0);
            }
            return null;
        }

        public static ValueVector CreateVector(RefEval re)
        {
            return new SheetVector(re);
        }
        private class StringLookupComparer : LookupValueComparerBase
        {

            protected readonly String _value;
            protected readonly Regex _wildCardPattern;
            protected readonly bool _matchExact;
            protected readonly bool _isMatchFunction;

            public StringLookupComparer(StringEval se, bool matchExact, bool isMatchFunction)
                : base(se)
            {

                _value = se.StringValue;
                _wildCardPattern = Countif.StringMatcher.GetWildCardPattern(_value);
                _matchExact = matchExact;
                _isMatchFunction = isMatchFunction;
            }

            protected virtual String ConvertToString(ValueEval other)
            {
                StringEval se = (StringEval)other;
                return se.StringValue;
            }

            protected override CompareResult CompareSameType(ValueEval other)
            {
                String stringValue = ConvertToString(other);
                if (_wildCardPattern != null && (_isMatchFunction || !_matchExact))
                {
                    MatchCollection matcher = _wildCardPattern.Matches(stringValue);
                    bool matches = matcher.Count > 0;
                    return CompareResult.ValueOf(matches);
                }

                return CompareResult.ValueOf(String.Compare(_value, stringValue, true));
            }

            protected override String GetValueAsString()
            {
                return _value;
            }
        }
        private sealed class TolerantStringLookupComparer : StringLookupComparer
        {
            static StringEval ConvertToStringEval(ValueEval eval)
            {
                if (eval is StringEval)
                {
                    return (StringEval)eval;
                }
                String sv = OperandResolver.CoerceValueToString(eval);
                return new StringEval(sv);
            }

            public TolerantStringLookupComparer(ValueEval eval, bool matchExact, bool isMatchFunction) :
                base(ConvertToStringEval(eval), matchExact, isMatchFunction)
            {

            }

            protected override String ConvertToString(ValueEval other)
            {
                return OperandResolver.CoerceValueToString(other);
            }
        }

        private sealed class NumberLookupComparer : LookupValueComparerBase
        {
            private readonly double _value;

            public NumberLookupComparer(NumberEval ne)
                : base(ne)
            {

                _value = ne.NumberValue;
            }
            protected override CompareResult CompareSameType(ValueEval other)
            {
                NumberEval ne = (NumberEval)other;
                return CompareResult.ValueOf(_value.CompareTo(ne.NumberValue));
            }
            protected override String GetValueAsString()
            {
                return _value.ToString(CultureInfo.InvariantCulture);
            }
        }


        /**
         * Processes the third argument to VLOOKUP, or HLOOKUP (<b>col_index_num</b> 
         * or <b>row_index_num</b> respectively).<br/>
         * Sample behaviour:
         *    <table border="0" cellpAdding="1" cellspacing="2" summary="Sample behaviour">
         *      <tr><th>Input Return</th><th>Value </th><th>Thrown Error</th></tr>
         *      <tr><td>5</td><td>4</td><td> </td></tr>
         *      <tr><td>2.9</td><td>2</td><td> </td></tr>
         *      <tr><td>"5"</td><td>4</td><td> </td></tr>
         *      <tr><td>"2.18e1"</td><td>21</td><td> </td></tr>
         *      <tr><td>"-$2"</td><td>-3</td><td>*</td></tr>
         *      <tr><td>FALSE</td><td>-1</td><td>*</td></tr>
         *      <tr><td>TRUE</td><td>0</td><td> </td></tr>
         *      <tr><td>"TRUE"</td><td> </td><td>#REF!</td></tr>
         *      <tr><td>"abc"</td><td> </td><td>#REF!</td></tr>
         *      <tr><td>""</td><td> </td><td>#REF!</td></tr>
         *      <tr><td>&lt;blank&gt;</td><td> </td><td>#VALUE!</td></tr>
         *    </table><br/>
         *    
         *  * Note - out of range errors (both too high and too low) are handled by the caller. 
         * @return column or row index as a zero-based value
         * 
         */
        public static int ResolveRowOrColIndexArg(ValueEval rowColIndexArg, int srcCellRow, int srcCellCol)
        {
            if (rowColIndexArg == null) {
                throw new ArgumentException("argument must not be null");
            }

            ValueEval veRowColIndexArg;
            try {
                veRowColIndexArg = OperandResolver.GetSingleValue(rowColIndexArg, srcCellRow, (short)srcCellCol);
            } catch (EvaluationException) {
                // All errors get translated to #REF!
                throw EvaluationException.InvalidRef();
            }
            int oneBasedIndex;
            if (veRowColIndexArg is StringEval) {
                StringEval se = (StringEval)veRowColIndexArg;
                String strVal = se.StringValue;
                Double dVal = OperandResolver.ParseDouble(strVal);
                if (Double.IsNaN(dVal)) {
                    // String does not resolve to a number. Raise #REF! error.
                    throw EvaluationException.InvalidRef();
                    // This includes text booleans "TRUE" and "FALSE".  They are not valid.
                }
                // else - numeric value parses OK
            }
            // actual BoolEval values get interpreted as FALSE->0 and TRUE->1
            oneBasedIndex = OperandResolver.CoerceValueToInt(veRowColIndexArg);
            if (oneBasedIndex < 1) {
                // note this is asymmetric with the errors when the index is too large (#REF!)  
                throw EvaluationException.InvalidValue();
            }
            return oneBasedIndex - 1; // convert to zero based
        }



        /**
         * The second argument (table_array) should be an area ref, but can actually be a cell ref, in
         * which case it Is interpreted as a 1x1 area ref.  Other scalar values cause #VALUE! error.
         */
        public static AreaEval ResolveTableArrayArg(ValueEval eval)
        {
            if (eval is AreaEval)
            {
                return (AreaEval)eval;
            }

            if (eval is RefEval) {
                RefEval refEval = (RefEval)eval;
                // Make this cell ref look like a 1x1 area ref.

                // It doesn't matter if eval is a 2D or 3D ref, because that detail is never asked of AreaEval.
                return refEval.Offset(0, 0, 0, 0);
            }
            throw EvaluationException.InvalidValue();
        }


        /**
         * Resolves the last (optional) parameter (<b>range_lookup</b>) to the VLOOKUP and HLOOKUP functions. 
         * @param rangeLookupArg
         * @param srcCellRow
         * @param srcCellCol
         * @return
         * @throws EvaluationException
         */
        public static bool ResolveRangeLookupArg(ValueEval rangeLookupArg, int srcCellRow, int srcCellCol)
        {
            ValueEval valEval = OperandResolver.GetSingleValue(rangeLookupArg, srcCellRow, srcCellCol);
            if (valEval == MissingArgEval.instance)
            {
                // Tricky:
                // forth arg exists but is not supplied: "=VLOOKUP(A1,A2:A4,2,)"
                return false;
            }
            if (valEval is BlankEval)
            {
                // Tricky:
                // fourth arg supplied but Evaluates to blank
                // this does not Get the default value
                return false;
            }
            if (valEval is BoolEval)
            {
                // Happy day flow 
                BoolEval boolEval = (BoolEval)valEval;
                return boolEval.BooleanValue;
            }

            if (valEval is StringEval)
            {
                String stringValue = ((StringEval)valEval).StringValue;
                if (stringValue.Length < 1)
                {
                    // More trickiness:
                    // Empty string Is not the same as BlankEval.  It causes #VALUE! error 
                    throw EvaluationException.InvalidValue();
                }
                // TODO move parseBoolean to OperandResolver
                bool? b = Countif.ParseBoolean(stringValue);
                if (b != null)
                {
                    // string Converted to bool OK
                    return b == true ? true : false;
                }
                //// Even more trickiness:
                //// Note - even if the StringEval represents a number value (for example "1"), 
                //// Excel does not resolve it to a bool.  
                throw EvaluationException.InvalidValue();
                //// This Is in contrast to the code below,, where NumberEvals values (for 
                //// example 0.01) *do* resolve to equivalent bool values.
            }
            if (valEval is NumericValueEval)
            {
                NumericValueEval nve = (NumericValueEval)valEval;
                // zero Is FALSE, everything else Is TRUE
                return 0.0 != nve.NumberValue;
            }
            throw new Exception("Unexpected eval type (" + valEval.GetType().Name + ")");
        }

        public static int lookupFirstIndexOfValue(ValueEval lookupValue, ValueVector vector, bool isRangeLookup)
        {
            LookupValueComparer lookupComparer = CreateLookupComparer(lookupValue, isRangeLookup, false);
            int result;
            if (isRangeLookup)
            {
                result = PerformBinarySearch(vector, lookupComparer);
            }
            else
            {
                result = lookupFirstIndexOfValue(lookupComparer, vector, MatchMode.ExactMatch);
            }
            if (result < 0)
            {
                throw new EvaluationException(ErrorEval.NA);
            }
            return result;
        }
        public static int XlookupIndexOfValue(ValueEval lookupValue, ValueVector vector, MatchMode matchMode, SearchMode searchMode)
        {
            LookupValueComparer lookupComparer = CreateTolerantLookupComparer(lookupValue, true, true);
            int result;
            if (searchMode == SearchMode.IterateBackward || searchMode == SearchMode.BinarySearchBackward)
            {
                result = lookupLastIndexOfValue(lookupComparer, vector, matchMode);
            }
            else
            {
                result = lookupFirstIndexOfValue(lookupComparer, vector, matchMode);
            }
            if (result < 0)
            {
                throw new EvaluationException(ErrorEval.NA);
            }
            return result;
        }

        /**
         * Finds first (lowest index) exact occurrence of specified value.
         * @param lookupComparer the value to be found in column or row vector
         * @param vector the values to be searched. For VLOOKUP this Is the first column of the 
         * 	tableArray. For HLOOKUP this Is the first row of the tableArray. 
         * @return zero based index into the vector, -1 if value cannot be found
         */
        private static int lookupFirstIndexOfValue(LookupValueComparer lookupComparer, ValueVector vector, MatchMode matchMode)
        {

            // Find first occurrence of lookup value
            int size = vector.Size;
            int bestMatchIdx = -1;
            ValueEval bestMatchEval = null;
            for (int i = 0; i < size; i++)
            {
                ValueEval valueEval = vector.GetItem(i);
                CompareResult result = lookupComparer.CompareTo(valueEval);
                if (result.IsEqual)
                {
                    return i;
                }
                switch (matchMode)
                {
                    case MatchMode.ExactMatchFallbackToLargerValue:
                        if (result.IsLessThan)
                        {
                            if (bestMatchEval == null)
                            {
                                bestMatchIdx = i;
                                bestMatchEval = valueEval;
                            }
                            else
                            {
                                LookupValueComparer matchComparer = CreateTolerantLookupComparer(valueEval, true, true);
                                if (matchComparer.CompareTo(bestMatchEval).IsLessThan)
                                {
                                    bestMatchIdx = i;
                                    bestMatchEval = valueEval;
                                }
                            }
                        }
                        break;
                    case MatchMode.ExactMatchFallbackToSmallerValue:
                        if (result.IsGreaterThan)
                        {
                            if (bestMatchEval == null)
                            {
                                bestMatchIdx = i;
                                bestMatchEval = valueEval;
                            }
                            else
                            {
                                LookupValueComparer matchComparer = CreateTolerantLookupComparer(valueEval, true, true);
                                if (matchComparer.CompareTo(bestMatchEval).IsGreaterThan)
                                {
                                    bestMatchIdx = i;
                                    bestMatchEval = valueEval;
                                }
                            }
                        }
                        break;
                }
            }
            return bestMatchIdx;
        }

        /**
 * Finds last (greatest index) matching occurrence of specified value.
 * @param lookupComparer the value to be found in column or row vector
 * @param vector the values to be searched. For VLOOKUP this is the first column of the
 *  tableArray. For HLOOKUP this is the first row of the tableArray.
 * @param matchMode
 * @return zero based index into the vector, -1 if value cannot be found
 */
        private static int lookupLastIndexOfValue(LookupValueComparer lookupComparer, ValueVector vector,
                                                  MatchMode matchMode)
        {
            // find last occurrence of lookup value
            int size = vector.Size;
            int bestMatchIdx = -1;
            ValueEval bestMatchEval = null;
            for (int i = size - 1; i >= 0; i--)
            {
                ValueEval valueEval = vector.GetItem(i);
                CompareResult result = lookupComparer.CompareTo(valueEval);
                if (result.IsEqual)
                {
                    return i;
                }
                switch (matchMode)
                {
                    case MatchMode.ExactMatchFallbackToLargerValue:
                        if (result.IsLessThan)
                        {
                            if (bestMatchEval == null)
                            {
                                bestMatchIdx = i;
                                bestMatchEval = valueEval;
                            }
                            else
                            {
                                LookupValueComparer matchComparer = CreateTolerantLookupComparer(valueEval, true, true);
                                if (matchComparer.CompareTo(bestMatchEval).IsLessThan)
                                {
                                    bestMatchIdx = i;
                                    bestMatchEval = valueEval;
                                }
                            }
                        }
                        break;
                    case MatchMode.ExactMatchFallbackToSmallerValue:
                        if (result.IsGreaterThan)
                        {
                            if (bestMatchEval == null)
                            {
                                bestMatchIdx = i;
                                bestMatchEval = valueEval;
                            }
                            else
                            {
                                LookupValueComparer matchComparer = CreateTolerantLookupComparer(valueEval, true, true);
                                if (matchComparer.CompareTo(bestMatchEval).IsGreaterThan)
                                {
                                    bestMatchIdx = i;
                                    bestMatchEval = valueEval;
                                }
                            }
                        }
                        break;
                }
            }
            return bestMatchIdx;
        }

        /**
         * Excel has funny behaviour when the some elements in the search vector are the wrong type.
         * 
         */
        private static int PerformBinarySearch(ValueVector vector, LookupValueComparer lookupComparer)
        {
            // both low and high indexes point to values assumed too low and too high.
            BinarySearchIndexes bsi = new BinarySearchIndexes(vector.Size);

            while (true)
            {
                int midIx = bsi.GetMidIx();

                if (midIx < 0)
                {
                    return bsi.GetLowIx();
                }
                CompareResult cr = lookupComparer.CompareTo(vector.GetItem(midIx));
                if (cr.IsTypeMismatch)
                {
                    int newMidIx = HandleMidValueTypeMismatch(lookupComparer, vector, bsi, midIx);
                    if (newMidIx < 0)
                    {
                        continue;
                    }
                    midIx = newMidIx;
                    cr = lookupComparer.CompareTo(vector.GetItem(midIx));
                }
                if (cr.IsEqual)
                {
                    return FindLastIndexInRunOfEqualValues(lookupComparer, vector, midIx, bsi.GetHighIx());
                }
                bsi.NarrowSearch(midIx, cr.IsLessThan);
            }
        }
        /**
         * Excel seems to handle mismatched types initially by just stepping 'mid' ix forward to the 
         * first compatible value.
         * @param midIx 'mid' index (value which has the wrong type)
         * @return usually -1, signifying that the BinarySearchIndex has been narrowed to the new mid 
         * index.  Zero or greater signifies that an exact match for the lookup value was found
         */
        private static int HandleMidValueTypeMismatch(LookupValueComparer lookupComparer, ValueVector vector,
                BinarySearchIndexes bsi, int midIx)
        {
            int newMid = midIx;
            int highIx = bsi.GetHighIx();

            while (true)
            {
                newMid++;
                if (newMid == highIx)
                {
                    // every element from midIx to highIx was the wrong type
                    // move highIx down to the low end of the mid values
                    bsi.NarrowSearch(midIx, true);
                    return -1;
                }
                CompareResult cr = lookupComparer.CompareTo(vector.GetItem(newMid));
                if (cr.IsLessThan && newMid == highIx - 1)
                {
                    // move highIx down to the low end of the mid values
                    bsi.NarrowSearch(midIx, true);
                    return -1;
                    // but only when "newMid == highIx-1"? slightly weird.
                    // It would seem more efficient to always do this.
                }
                if (cr.IsTypeMismatch)
                {
                    // keep stepping over values Until the right type Is found
                    continue;
                }
                if (cr.IsEqual)
                {
                    return newMid;
                }
                // Note - if moving highIx down (due to lookup<vector[newMid]),
                // this execution path only moves highIx it down as far as newMid, not midIx,
                // which would be more efficient.
                bsi.NarrowSearch(newMid, cr.IsLessThan);
                return -1;
            }
        }
        /**
         * Once the binary search has found a single match, (V/H)LOOKUP steps one by one over subsequent
         * values to choose the last matching item.
         */
        private static int FindLastIndexInRunOfEqualValues(LookupValueComparer lookupComparer, ValueVector vector,
                    int firstFoundIndex, int maxIx)
        {
            for (int i = firstFoundIndex + 1; i < maxIx; i++)
            {
                if (!lookupComparer.CompareTo(vector.GetItem(i)).IsEqual)
                {
                    return i - 1;
                }
            }
            return maxIx - 1;
        }

        public static LookupValueComparer CreateLookupComparer(ValueEval lookupValue, bool matchExact, bool isMatchFunction)
        {

            if (lookupValue == BlankEval.instance)
            {
                // blank eval translates to zero
                // Note - a blank eval in the lookup column/row never matches anything
                // empty string in the lookup column/row can only be matched by explicit emtpty string
                return new NumberLookupComparer(NumberEval.ZERO);
            }
            if (lookupValue is StringEval)
            {
                return new StringLookupComparer((StringEval)lookupValue, matchExact, isMatchFunction);
            }
            if (lookupValue is NumberEval)
            {
                return new NumberLookupComparer((NumberEval)lookupValue);
            }
            if (lookupValue is BoolEval)
            {
                return new BooleanLookupComparer((BoolEval)lookupValue);
            }
            throw new ArgumentException("Bad lookup value type (" + lookupValue.GetType().Name + ")");
        }
        private static LookupValueComparer CreateTolerantLookupComparer(ValueEval lookupValue, bool matchExact, bool isMatchFunction)
        {
            if (lookupValue == BlankEval.instance)
            {
                return new TolerantStringLookupComparer(new StringEval(""), matchExact, isMatchFunction);
            }
            if (lookupValue is BoolEval) {
                return new BooleanLookupComparer((BoolEval)lookupValue);
            }
            if (matchExact && lookupValue is NumberEval) {
                return new NumberLookupComparer((NumberEval)lookupValue);
            }
            return new TolerantStringLookupComparer(lookupValue, matchExact, isMatchFunction);
        }
    }
    /**
    * Enumeration to support <b>4</b> valued comparison results.<p/>
    * Excel lookup functions have complex behaviour in the case where the lookup array has mixed 
    * types, and/or Is Unordered.  Contrary to suggestions in some Excel documentation, there
    * does not appear to be a Universal ordering across types.  The binary search algorithm used
    * Changes behaviour when the Evaluated 'mid' value has a different type to the lookup value.<p/>
    * 
    * A simple int might have done the same job, but there Is risk in confusion with the well 
    * known <c>Comparable.CompareTo()</c> and <c>Comparator.Compare()</c> which both use
    * a ubiquitous 3 value result encoding.
    */
    public class CompareResult
    {
        private readonly bool _isTypeMismatch;
        private readonly bool _isLessThan;
        private readonly bool _isEqual;
        private readonly bool _isGreaterThan;

        private CompareResult(bool IsTypeMismatch, int simpleCompareResult)
        {
            if (IsTypeMismatch)
            {
                _isTypeMismatch = true;
                _isLessThan = false;
                _isEqual = false;
                _isGreaterThan = false;
            }
            else
            {
                _isTypeMismatch = false;
                _isLessThan = simpleCompareResult < 0;
                _isEqual = simpleCompareResult == 0;
                _isGreaterThan = simpleCompareResult > 0;
            }
        }

        public static readonly CompareResult TypeMismatch = new CompareResult(true, 0);
        public static readonly CompareResult LessThan = new CompareResult(false, -1);
        public static readonly CompareResult Equal = new CompareResult(false, 0);
        public static readonly CompareResult GreaterThan = new CompareResult(false, +1);

        public static CompareResult ValueOf(int simpleCompareResult)
        {
            if (simpleCompareResult < 0)
            {
                return LessThan;
            }
            if (simpleCompareResult > 0)
            {
                return GreaterThan;
            }
            return Equal;
        }
        public static CompareResult ValueOf(bool matches)
        {
            if (matches)
            {
                return Equal;
            }
            return LessThan;
        }
        public bool IsTypeMismatch
        {
            get { return _isTypeMismatch; }
        }
        public bool IsLessThan
        {
            get { return _isLessThan; }
        }
        public bool IsEqual
        {
            get { return _isEqual; }
        }
        public bool IsGreaterThan
        {
            get { return _isGreaterThan; }
        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append(FormatAsString);
            sb.Append("]");
            return sb.ToString();
        }

        private String FormatAsString
        {
            get
            {
                if (_isTypeMismatch)
                {
                    return "TYPE_MISMATCH";
                }
                if (_isLessThan)
                {
                    return "LESS_THAN";
                }
                if (_isEqual)
                {
                    return "EQUAL";
                }
                if (_isGreaterThan)
                {
                    return "GREATER_THAN";
                }
                // toString must be reliable
                return "error";
            }
        }
    }
    /**
    * Encapsulates some standard binary search functionality so the Unusual Excel behaviour can
    * be clearly distinguished. 
    */
    internal sealed class BinarySearchIndexes
    {

        private int _lowIx;
        private int _highIx;

        public BinarySearchIndexes(int highIx)
        {
            _lowIx = -1;
            _highIx = highIx;
        }

        /**
         * @return -1 if the search range Is empty
         */
        public int GetMidIx()
        {
            int ixDiff = _highIx - _lowIx;
            if (ixDiff < 2)
            {
                return -1;
            }
            return _lowIx + (ixDiff / 2);
        }

        public int GetLowIx()
        {
            return _lowIx;
        }
        public int GetHighIx()
        {
            return _highIx;
        }
        public void NarrowSearch(int midIx, bool isLessThan)
        {
            if (isLessThan)
            {
                _highIx = midIx;
            }
            else
            {
                _lowIx = midIx;
            }
        }
    }
    internal sealed class BooleanLookupComparer : LookupValueComparerBase
    {
        private readonly bool _value;

        public BooleanLookupComparer(BoolEval be)
            : base(be)
        {

            _value = be.BooleanValue;
        }
        protected override CompareResult CompareSameType(ValueEval other)
        {
            BoolEval be = (BoolEval)other;
            bool otherVal = be.BooleanValue;
            if (_value == otherVal)
            {
                return CompareResult.Equal;
            }
            // TRUE > FALSE
            if (_value)
            {
                return CompareResult.GreaterThan;
            }
            return CompareResult.LessThan;
        }
        protected override String GetValueAsString()
        {
            return _value.ToString();
        }
    }
    /**
    * Represents a single row or column within an <c>AreaEval</c>.
    */
    public interface ValueVector
    {
        ValueEval GetItem(int index);
        int Size { get; }
    }


    public interface LookupValueComparer
    {
        /**
         * @return one of 4 instances or <c>CompareResult</c>: <c>LESS_THAN</c>, <c>EQUAL</c>, 
         * <c>GREATER_THAN</c> or <c>TYPE_MISMATCH</c>
         */
        CompareResult CompareTo(ValueEval other);
    }



    internal abstract class LookupValueComparerBase : LookupValueComparer
    {

        private readonly Type _targetType;
        protected LookupValueComparerBase(ValueEval targetValue)
        {
            if (targetValue == null)
            {
                throw new Exception("targetValue cannot be null");
            }
            _targetType = targetValue.GetType();
        }
        public CompareResult CompareTo(ValueEval other)
        {
            if (other == null)
            {
                throw new Exception("Compare to value cannot be null");
            }
            if (_targetType != other.GetType())
            {
                return CompareResult.TypeMismatch;
            }
            return CompareSameType(other);
        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append(GetValueAsString());
            sb.Append("]");
            return sb.ToString();
        }
        protected abstract CompareResult CompareSameType(ValueEval other);
        /** used only for debug purposes */
        protected abstract String GetValueAsString();
    }

}
