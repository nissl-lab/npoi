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
    using NPOI.SS.Formula.Eval;


    public class SimpleValueVector : ValueVector
    {
        private ValueEval[] _values;

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
    public class Lookup : Function
    {


        public ValueEval Evaluate(ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            switch (args.Length)
            {
                case 3:
                    break;
                case 2:
                    // complex rules to choose lookupVector and resultVector from the single area ref
                    throw new Exception("Two arg version of LOOKUP not supported yet");
                default:
                    return ErrorEval.VALUE_INVALID;
            }


            try
            {
                ValueEval lookupValue = OperandResolver.GetSingleValue(args[0], srcCellRow, srcCellCol);
                AreaEval aeLookupVector = LookupUtils.ResolveTableArrayArg(args[1]);
                AreaEval aeResultVector = LookupUtils.ResolveTableArrayArg(args[2]);

                ValueVector lookupVector = CreateVector(aeLookupVector);
                ValueVector resultVector = CreateVector(aeResultVector);
                if (lookupVector.Size > resultVector.Size)
                {
                    // Excel seems to handle this by accessing past the end of the result vector.
                    throw new Exception("Lookup vector and result vector of differing sizes not supported yet");
                }
                int index = LookupUtils.LookupIndexOfValue(lookupValue, lookupVector, true);

                return resultVector.GetItem(index);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        private static ValueVector CreateVector(AreaEval ae)
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
