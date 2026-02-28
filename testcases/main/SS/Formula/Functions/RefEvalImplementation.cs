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

namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;

    public class RefEvalImplementation : RefEval
    {

        private ValueEval value;

        public RefEvalImplementation(ValueEval value)
        {
            this.value = value;
        }


        public AreaEval Offset(int relFirstRowIx, int relLastRowIx,
                int relFirstColIx, int relLastColIx)
        {
            throw new InvalidOperationException();
        }


        public ValueEval GetInnerValueEval(int sheetIndex)
        {
            return value;
        }


        public int NumberOfSheets
        {
            get
            {
                return 1;
            }
        }


        public int FirstSheetIndex
        {
            get
            {
                return 0;
            }
        }


        public int LastSheetIndex
        {
            get
            {
                return 0;
            }
        }


        public int Row
        {
            get
            {
                throw new InvalidOperationException();
            }
        }


        public int Column
        {
            get
            {
                throw new InvalidOperationException();
            }
        }

    }

}