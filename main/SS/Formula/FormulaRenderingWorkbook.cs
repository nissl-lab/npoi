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
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;

    /**
     * Abstracts a workbook for the purpose of converting formula To text.<br/>
     * 
     * For POI internal use only
     * 
     * @author Josh Micich
     */
    public interface IFormulaRenderingWorkbook
    {

        /**
         * @return <c>null</c> if externSheetIndex refers To a sheet inside the current workbook
         */
        ExternalSheet GetExternalSheet(int externSheetIndex);
        //String GetSheetNameByExternSheet(int externSheetIndex);
        /**
         * @return the name of the (first) sheet referred to by the given external sheet index
         */
        String GetSheetFirstNameByExternSheet(int externSheetIndex);
        /**
         * @return the name of the (last) sheet referred to by the given external sheet index
         */
        String GetSheetLastNameByExternSheet(int externSheetIndex);
        String ResolveNameXText(NameXPtg nameXPtg);
        String GetNameText(NamePtg namePtg);
    }
}