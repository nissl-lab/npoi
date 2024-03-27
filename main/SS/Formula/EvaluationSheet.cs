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

    /// <summary>
    /// <para>
    /// Abstracts a sheet for the purpose of formula evaluation.<br/>
    /// </para>
    /// <para>
    /// For POI internal use only
    /// </para>
    /// </summary>
    /// @author Josh Micich
    public interface IEvaluationSheet
    {

        /// <summary>
        /// </summary>
        /// <returns><c>null</c> if there is no cell at the specified coordinates</returns>
        IEvaluationCell GetCell(int rowIndex, int columnIndex);
        /// <summary>
        /// Propagated from <see cref="EvaluationWorkbook.clearAllCachedResultValues()" /> to clear locally cached data.
        /// </summary>
        /// <see cref="WorkbookEvaluator.clearAllCachedResultValues()" />
        /// <see cref="EvaluationWorkbook.clearAllCachedResultValues()" />
        void ClearAllCachedResultValues();
    }
}