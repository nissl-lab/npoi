/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI.SS.Formula.Eval
{
    using System;

    using NPOI.SS.Formula;
    using System.Text;

    /**
     * Evaluation of a Name defined in a Sheet or Workbook scope
     */
    public class ExternalNameEval : ValueEval
    {
        private IEvaluationName _name;

        public ExternalNameEval(IEvaluationName name)
        {
            _name = name;
        }

        public IEvaluationName Name
        {
            get
            {
                return _name;
            }
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append(_name.NameText);
            sb.Append("]");
            return sb.ToString();
        }
    }

}