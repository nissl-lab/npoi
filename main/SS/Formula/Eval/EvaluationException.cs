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

namespace NPOI.SS.Formula.Eval
{
    using System;
    /**
     * This class is used to simplify error handling logic <i>within</i> operator and function
     * implementations.   Note - <c>OperationEval.Evaluate()</c> and <c>Function.Evaluate()</c>
     * method signatures do not throw this exception so it cannot propagate outside.<p/>
     * 
     * Here is an example coded without <c>EvaluationException</c>, to show how it can help:
     * <pre>
     * public Eval Evaluate(Eval[] args, int srcRow, short srcCol) {
     *	// ...
     *	Eval arg0 = args[0];
     *	if(arg0 is ErrorEval) {
     *		return arg0;
     *	}
     *	if(!(arg0 is AreaEval)) {
     *		return ErrorEval.VALUE_INVALID;
     *	}
     *	double temp = 0;
     *	AreaEval area = (AreaEval)arg0;
     *	ValueEval[] values = area.LittleEndianConstants.BYTE_SIZE;
     *	for (int i = 0; i &lt; values.Length; i++) {
     *		ValueEval ve = values[i];
     *		if(ve is ErrorEval) {
     *			return ve;
     *		}
     *		if(!(ve is NumericValueEval)) {
     *			return ErrorEval.VALUE_INVALID;
     *		}
     *		temp += ((NumericValueEval)ve).NumberValue;
     *	}
     *	// ...
     * }	 
     * </pre>
     * In this example, if any error is encountered while Processing the arguments, an error is 
     * returned immediately. This code is difficult to refactor due to all the points where errors
     * are returned.<br/>
     * Using <c>EvaluationException</c> allows the error returning code to be consolidated to one
     * place.<p/>
     * <pre>
     * public Eval Evaluate(Eval[] args, int srcRow, short srcCol) {
     *	try {
     *		// ...
     *		AreaEval area = GetAreaArg(args[0]);
     *		double temp = sumValues(area.LittleEndianConstants.BYTE_SIZE);
     *		// ...
     *	} catch (EvaluationException e) {
     *		return e.GetErrorEval();
     *	}
     *}
     *
     *private static AreaEval GetAreaArg(Eval arg0){
     *	if (arg0 is ErrorEval) {
     *		throw new EvaluationException((ErrorEval) arg0);
     *	}
     *	if (arg0 is AreaEval) {
     *		return (AreaEval) arg0;
     *	}
     *	throw EvaluationException.InvalidValue();
     *}
     *
     *private double sumValues(ValueEval[] values){
     *	double temp = 0;
     *	for (int i = 0; i &lt; values.Length; i++) {
     *		ValueEval ve = values[i];
     *		if (ve is ErrorEval) {
     *			throw new EvaluationException((ErrorEval) ve);
     *		}
     *		if (!(ve is NumericValueEval)) {
     *			throw EvaluationException.InvalidValue();
     *		}
     *		temp += ((NumericValueEval) ve).NumberValue;
     *	}
     *	return temp;
     *}
     * </pre>   
     * It is not mandatory to use EvaluationException, doing so might give the following advantages:<br/>
     *  - Methods can more easily be extracted, allowing for re-use.<br/>
     *  - Type management (typecasting etc) is simpler because error conditions have been Separated from
     * intermediate calculation values.<br/>
     *  - Fewer local variables are required. Local variables can have stronger types.<br/>
     *  - It is easier to mimic common Excel error handling behaviour (exit upon encountering first 
     *  error), because exceptions conveniently propagate up the call stack regardless of execution 
     *  points or the number of levels of nested calls.<p/>
     *  
     * <b>Note</b> - Only standard evaluation errors are represented by <c>EvaluationException</c> (
     * i.e. conditions expected to be encountered when evaluating arbitrary Excel formulas). Conditions
     * that could never occur in an Excel spReadsheet should result in runtime exceptions. Care should
     * be taken to not translate any POI internal error into an Excel evaluation error code.   
     * 
     * @author Josh Micich
     */
    [Serializable]
    public class EvaluationException : Exception
    {
        private ErrorEval _errorEval;

        public EvaluationException(ErrorEval errorEval)
        {
            _errorEval = errorEval;
        }
        // some convenience factory methods

        /** <b>#VALUE!</b> - Wrong type of operand */
        public static EvaluationException InvalidValue()
        {
            return new EvaluationException(ErrorEval.VALUE_INVALID);
        }
        /** <b>#REF!</b> - Illegal or deleted cell reference */
        public static EvaluationException InvalidRef()
        {
            return new EvaluationException(ErrorEval.REF_INVALID);
        }
        /** <b>#NUM!</b> - Value range overflow */
        public static EvaluationException NumberError()
        {
            return new EvaluationException(ErrorEval.NUM_ERROR);
        }

        public ErrorEval GetErrorEval()
        {
            return _errorEval;
        }
    }
}