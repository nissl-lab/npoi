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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    using ExtendedNumerics;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Util;
    /// <summary>
    /// Implementation of the DStdevp function:
    /// Gets the standard deviation value of a column in an area with given conditions.
    /// </summary>
    public sealed class DStdevp : IDStarAlgorithm
    {
        private  List<NumericValueEval> values = new List<NumericValueEval>();
        public bool ProcessMatch(ValueEval eval)
        {
            if(eval is NumericValueEval valueEval)
            {
                values.Add(valueEval);
            }
            return true;
        }
        public ValueEval Result
        {
            get
            {
                double[] array = new double[values.Count];
                int pos = 0;
                foreach(NumericValueEval d in values)
                {
                    array[pos++] = d.NumberValue;
                }
                double stdev = StatsLib.stdevp(array);
                return new NumberEval((double)BigDecimal.Parse(NumberToTextConverter.ToText(stdev)));
            }
            
        }

        public bool AllowEmptyMatchField { get; } = false;
    }
}
