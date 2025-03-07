/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using System;

namespace NPOI.SS.Formula
{
    internal class SingleValueAttribute : Attribute
    {
        private readonly bool _isSingleValue=false;
        public SingleValueAttribute(bool isSingleValue)
        {
            this._isSingleValue = isSingleValue;
        }
        public bool IsSingleValue
        {
            get { return _isSingleValue; }
        }
    }

    /// <summary>
    /// Enumeration of various formula types. For internal use only
    /// </summary>
    public enum FormulaType:int
    {
        [SingleValue(true)]
        Cell = 0,
        [SingleValue(true)]
        Shared = 1,
        [SingleValue(false)]
        Array = 2,
        [SingleValue(true)]
        CondFormat = 3,
        [SingleValue(false)]
        NamedRange = 4,
        // this constant is currently very specific.  The exact differences from general data
        // validation formulas or conditional format formulas is not known yet
        [SingleValue(false)]
        DataValidationList = 5,
    }
}