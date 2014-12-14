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
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula;
    /**
     * Implementation of the VLOOKUP() function.<p/>
     * 
     * VLOOKUP Finds a row in a lookup table by the first column value and returns the value from another column.<br/>
     * 
     * <b>Syntax</b>:<br/>
     * <b>VLOOKUP</b>(<b>lookup_value</b>, <b>table_array</b>, <b>col_index_num</b>, range_lookup)<p/>
     * 
     * <b>lookup_value</b>  The value to be found in the first column of the table array.<br/>
     * <b>table_array</b> An area reference for the lookup data. <br/>
     * <b>col_index_num</b> a 1 based index specifying which column value of the lookup data will be returned.<br/>
     * <b>range_lookup</b> If TRUE (default), VLOOKUP Finds the largest value less than or equal to 
     * the lookup_value.  If FALSE, only exact Matches will be considered<br/>   
     * 
     * @author Josh Micich
     */
    public class Vlookup : Var3or4ArgFunction
    {

        //private class ColumnVector : ValueVector
        //{

        //    private AreaEval _tableArray;
        //    private int _size;
        //    private int _columnAbsoluteIndex;
        //    private int _firstRowAbsoluteIndex;

        //    public ColumnVector(AreaEval tableArray, int columnIndex)
        //    {
        //        _columnAbsoluteIndex = tableArray.FirstColumn + columnIndex;
        //        if (!tableArray.ContainsColumn((short)_columnAbsoluteIndex))
        //        {
        //            int lastColIx = tableArray.LastColumn - tableArray.FirstColumn;
        //            throw new ArgumentException("Specified column index (" + columnIndex
        //                    + ") Is outside the allowed range (0.." + lastColIx + ")");
        //        }
        //        _tableArray = tableArray;
        //        _size = tableArray.LastRow - tableArray.FirstRow + 1;
        //        if (_size < 1)
        //        {
        //            throw new Exception("bad table array size zero");
        //        }
        //        _firstRowAbsoluteIndex = tableArray.FirstRow;
        //    }

        //    public ValueEval GetItem(int index)
        //    {
        //        if (index > _size)
        //        {
        //            throw new IndexOutOfRangeException("Specified index (" + index
        //                    + ") Is outside the allowed range (0.." + (_size - 1) + ")");
        //        }
        //        return _tableArray.GetValueAt(_firstRowAbsoluteIndex + index, (short)_columnAbsoluteIndex);
        //    }
        //    public int Size
        //    {
        //        get
        //        {
        //            return _size;
        //        }
        //    }
        //}

        private static ValueEval DEFAULT_ARG3 = BoolEval.TRUE;

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
        ValueEval arg2)
        {
            return Evaluate(srcRowIndex, srcColumnIndex, arg0, arg1, arg2, DEFAULT_ARG3);
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval lookup_value, ValueEval table_array,
                ValueEval col_index, ValueEval range_lookup)
        {
            try
            {
                // Evaluation order:
                // arg0 lookup_value, arg1 table_array, arg3 range_lookup, find lookup value, arg2 col_index, fetch result
                ValueEval lookupValue = OperandResolver.GetSingleValue(lookup_value, srcRowIndex, srcColumnIndex);
                TwoDEval tableArray = LookupUtils.ResolveTableArrayArg(table_array);
                bool isRangeLookup = LookupUtils.ResolveRangeLookupArg(range_lookup, srcRowIndex, srcColumnIndex);
                int rowIndex = LookupUtils.LookupIndexOfValue(lookupValue, LookupUtils.CreateColumnVector(tableArray, 0), isRangeLookup);
                int colIndex = LookupUtils.ResolveRowOrColIndexArg(col_index, srcRowIndex, srcColumnIndex);
                ValueVector resultCol = CreateResultColumnVector(tableArray, colIndex);
                return resultCol.GetItem(rowIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }


        /**
         * Returns one column from an <c>AreaEval</c>
         * 
         * @(#VALUE!) if colIndex Is negative, (#REF!) if colIndex Is too high
         */
        private ValueVector CreateResultColumnVector(TwoDEval tableArray, int colIndex)
        {
            if (colIndex >= tableArray.Width)
            {
                throw EvaluationException.InvalidRef();
            }
            return LookupUtils.CreateColumnVector(tableArray, colIndex);
        }
    }
}