
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
    using NPOI.SS.Formula.Functions;
    /**
     * Implementation of the HLOOKUP() function.<p/>
     * 
     * HLOOKUP Finds a column in a lookup table by the first row value and returns the value from another row.<br/>
     * 
     * <b>Syntax</b>:<br/>
     * <b>HLOOKUP</b>(<b>lookup_value</b>, <b>table_array</b>, <b>row_index_num</b>, range_lookup)<p/>
     * 
     * <b>lookup_value</b>  The value to be found in the first column of the table array.<br/>
     * <b>table_array</b> An area reference for the lookup data. <br/>
     * <b>row_index_num</b> a 1 based index specifying which row value of the lookup data will be returned.<br/>
     * <b>range_lookup</b> If TRUE (default), HLOOKUP Finds the largest value less than or equal to 
     * the lookup_value.  If FALSE, only exact Matches will be considered<br/>   
     * 
     * @author Josh Micich
     */
    public class Hlookup : Function
    {
        public ValueEval Evaluate(ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            ValueEval arg3 = null;
            switch (args.Length)
            {
                case 4:
                    arg3 = args[3]; // important: assumed array element Is never null
                    break;
                case 3:
                    break;
                default:
                    // wrong number of arguments
                    return ErrorEval.VALUE_INVALID;
            }
            try
            {
                // Evaluation order:
                // arg0 lookup_value, arg1 table_array, arg3 range_lookup, Find lookup value, arg2 row_index, fetch result
                ValueEval lookupValue = OperandResolver.GetSingleValue(args[0], srcCellRow, srcCellCol);
                AreaEval tableArray = LookupUtils.ResolveTableArrayArg(args[1]);
                bool IsRangeLookup = LookupUtils.ResolveRangeLookupArg(arg3, srcCellRow, srcCellCol);
                int colIndex = LookupUtils.LookupIndexOfValue(lookupValue, LookupUtils.CreateRowVector(tableArray, 0), IsRangeLookup);
                int rowIndex = LookupUtils.ResolveRowOrColIndexArg(args[2], srcCellRow, srcCellCol);
                ValueVector resultCol = CreateResultColumnVector(tableArray, rowIndex);
                return resultCol.GetItem(colIndex);
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
        private ValueVector CreateResultColumnVector(AreaEval tableArray, int rowIndex)
        {
            if (rowIndex >= tableArray.Height)
            {
                throw EvaluationException.InvalidRef();
            }
            return LookupUtils.CreateRowVector(tableArray, rowIndex);
        }
    }
}