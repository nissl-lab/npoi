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
    using NPOI.SS.Formula.Eval;

    /// <summary>
    /// Implementation of the DCount function:
    /// Counts the number of numeric cells in a column in an area with given conditions.
    /// </summary>
    public sealed class DCount : IDStarAlgorithm
    {
        private long count;
        public bool ProcessMatch(ValueEval eval)
        {
            if(eval is NumericValueEval)
            {
                count++;
            }
            return true;
        }
        public ValueEval Result
        {
            get
            {
                return new NumberEval(count);
            }
        }
        public bool AllowEmptyMatchField { get; } = true;
    }
}

