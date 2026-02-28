﻿/* ====================================================================
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

using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Eval
{
    /// <summary>
    /// Handling of a list of values, e.g. the 2nd argument in RANK(A1,(B1,B2,B3),1)
    /// </summary>
    public class RefListEval:ValueEval
    {
        private  List<ValueEval> list = new List<ValueEval>();

        public RefListEval(ValueEval v1, ValueEval v2)
        {
            Add(v1);
            Add(v2);
        }

        private void Add(ValueEval v)
        {
            // flatten multiple nested RefListEval
            if (v is RefListEval eval) {
                list.AddRange(eval.list);
            } else
            {
                list.Add(v);
            }
        }

        public List<ValueEval> GetList()
        {
            return list;
        }
    }
}
