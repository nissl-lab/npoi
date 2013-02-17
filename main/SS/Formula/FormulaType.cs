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

namespace NPOI.SS.Formula
{

    /**
     * Enumeration of various formula types.<br/>
     * 
     * For POI internal use only
     * 
     * @author Josh Micich
     */
    public enum FormulaType:int
    {
        Cell = 0,
        Shared = 1,
        Array = 2,
        CondFormat = 3,
        NamedRange = 4,
        // this constant is currently very specific.  The exact differences from general data
        // validation formulas or conditional format formulas is not known yet
        DataValidationList = 5,
    }
}