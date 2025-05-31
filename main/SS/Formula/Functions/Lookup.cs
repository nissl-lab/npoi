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

using NPOI.SS.Formula.Functions;

namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;


    public class SimpleValueVector : ValueVector
    {
        private readonly ValueEval[] _values;

        public SimpleValueVector(ValueEval[] values)
        {
            _values = values;
        }
        public ValueEval GetItem(int index)
        {
            return _values[index];
        }
        public int Size
        {
            get
            {
                return _values.Length;
            }
        }
    }
    /**
     * Implementation of Excel function LOOKUP.<p/>
     * 
     * LOOKUP Finds an index  row in a lookup table by the first column value and returns the value from another column.
     * 
     * <b>Syntax</b>:<br/>
     * <b>VLOOKUP</b>(<b>lookup_value</b>, <b>lookup_vector</b>, result_vector)<p/>
     * 
     * <b>lookup_value</b>  The value to be found in the lookup vector.<br/>
     * <b>lookup_vector</b> An area reference for the lookup data. <br/>
     * <b>result_vector</b> Single row or single column area reference from which the result value Is chosen.<br/>
     * 
     * @author Josh Micich
     */
    public class Lookup : Var2or3ArgFunction
    {


        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            // complex rules to choose lookupVector and resultVector from the single area ref

            try
            {
                /*
                The array form of LOOKUP is very similar to the HLOOKUP and VLOOKUP functions. The difference is that HLOOKUP searches for the value of lookup_value in the first row, VLOOKUP searches in the first column, and LOOKUP searches according to the dimensions of array.
                If array covers an area that is wider than it is tall (more columns than rows), LOOKUP searches for the value of lookup_value in the first row.
                If an array is square or is taller than it is wide (more rows than columns), LOOKUP searches in the first column.
                With the HLOOKUP and VLOOKUP functions, you can index down or across, but LOOKUP always selects the last value in the row or column.
                 */
                ValueEval lookupValue = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
                TwoDEval lookupArray = LookupUtils.ResolveTableArrayArg(arg1);
                ValueVector lookupVector;
                ValueVector resultVector;

                if (lookupArray.Width > lookupArray.Height)
                {
                    // If array covers an area that is wider than it is tall (more columns than rows), LOOKUP searches for the value of lookup_value in the first row.
                    lookupVector = CreateVector(lookupArray.GetRow(0));
                    resultVector = CreateVector(lookupArray.GetRow(lookupArray.Height - 1));
                }
                else
                {
                    // If an array is square or is taller than it is wide (more rows than columns), LOOKUP searches in the first column.
                    lookupVector = CreateVector(lookupArray.GetColumn(0));
                    resultVector = CreateVector(lookupArray.GetColumn(lookupArray.Width - 1));
                }
                // if a rectangular area reference was passed in as arg1, lookupVector and resultVector should be the same size
                //assert(lookupVector.getSize() == resultVector.getSize());

                int index = LookupUtils.LookupIndexOfValue(lookupValue, lookupVector, true);
                return resultVector.GetItem(index);
            }
            catch (EvaluationException e) {
                return e.GetErrorEval();
            }

        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
            ValueEval arg2)
        {
            try
            {
                ValueEval lookupValue = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
                TwoDEval aeLookupVector = LookupUtils.ResolveTableArrayArg(arg1);
                TwoDEval aeResultVector = LookupUtils.ResolveTableArrayArg(arg2);

                ValueVector lookupVector = CreateVector(aeLookupVector);
                ValueVector resultVector = CreateVector(aeResultVector);
                if (lookupVector.Size > resultVector.Size)
                {
                    // Excel seems to handle this by accessing past the end of the result vector.
                    throw new NPOI.Util.RuntimeException("Lookup vector and result vector of differing sizes not supported yet");
                }
                int index = LookupUtils.LookupIndexOfValue(lookupValue, lookupVector, true);

                return resultVector.GetItem(index);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
        private static ValueVector CreateVector(TwoDEval ae)
        {
            ValueVector result = LookupUtils.CreateVector(ae);
            if (result != null)
            {
                return result;
            }
            // extra complexity required to emulate the way LOOKUP can handles these abnormal cases.
            throw new InvalidOperationException("non-vector lookup or result areas not supported yet");
        }
    }
}
