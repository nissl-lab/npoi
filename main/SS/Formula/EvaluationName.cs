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

    using System;
    using NPOI.SS.Formula.PTG;
    /**
     * Abstracts a name record for formula evaluation.<br/>
     * 
     * For POI internal use only
     * 
     * @author Josh Micich
     */
    public interface IEvaluationName
    {

        String NameText{ get; }

        bool IsFunctionName{ get; }

        bool HasFormula { get; }

        Ptg[] NameDefinition { get; }

        bool IsRange{ get; }
        NamePtg CreatePtg();
    }
}

