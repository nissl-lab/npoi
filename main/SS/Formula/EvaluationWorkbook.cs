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
    using NPOI.SS.Formula.UDF;
    using NPOI.SS.Formula.PTG;

    public class ExternalSheet
    {
        private String _workbookName;
        private String _sheetName;

        public ExternalSheet(String workbookName, String sheetName)
        {
            _workbookName = workbookName;
            _sheetName = sheetName;
        }
        public String WorkbookName
        {
            get
            {
                return _workbookName;
            }
        }
        public String SheetName
        {
            get
            {
                return _sheetName;
            }
        }
    }

    public class ExternalSheetRange : ExternalSheet
    {
        private String _lastSheetName;
        public ExternalSheetRange(String workbookName, String firstSheetName, String lastSheetName)
            : base(workbookName, firstSheetName)
        {
            this._lastSheetName = lastSheetName;
        }

        public String FirstSheetName
        {
            get
            {
                return SheetName;
            }
        }
        public String LastSheetName
        {
            get
            {
                return _lastSheetName;
            }
        }
    }

    /// <summary>
    /// <para>
    /// Abstracts a workbook for the purpose of formula evaluation.<br/>
    /// </para>
    /// <para>
    /// For POI internal use only
    /// </para>
    /// </summary>
    /// @author Josh Micich
    public interface IEvaluationWorkbook
    {
        String GetSheetName(int sheetIndex);
        /// <summary>
        /// </summary>
        /// <returns>-1 if the specified sheet is from a different book</returns>
        int GetSheetIndex(IEvaluationSheet sheet);
        int GetSheetIndex(String sheetName);

        IEvaluationSheet GetSheet(int sheetIndex);

        /// <summary>
        /// <para>
        /// HSSF Only - fetch the external-style sheet details
        /// </para>
        /// <para>
        /// Return will have no workbook set if it's actually in our own workbook
        /// </para>
        /// </summary>
        ExternalSheet GetExternalSheet(int externSheetIndex);
        /// <summary>
        /// <para>
        /// XSSF Only - fetch the external-style sheet details
        /// </para>
        /// <para>
        /// Return will have no workbook set if it's actually in our own workbook
        /// </para>
        /// </summary>
        ExternalSheet GetExternalSheet(String firstSheetName, string lastSheetName, int externalWorkbookNumber);
        /// <summary>
        /// HSSF Only - convert an external sheet index to an internal sheet index,
        ///  for an external-style reference to one of this workbook's own sheets
        /// </summary>
        int ConvertFromExternSheetIndex(int externSheetIndex);
        /// <summary>
        /// HSSF Only - fetch the external-style name details
        /// </summary>
        ExternalName GetExternalName(int externSheetIndex, int externNameIndex);
        /// <summary>
        /// XSSF Only - fetch the external-style name details
        /// </summary>
        ExternalName GetExternalName(String nameName, String sheetName, int externalWorkbookNumber);

        IEvaluationName GetName(NamePtg namePtg);
        IEvaluationName GetName(String name, int sheetIndex);
        String ResolveNameXText(NameXPtg ptg);
        Ptg[] GetFormulaTokens(IEvaluationCell cell);
        UDFFinder GetUDFFinder();

        /// <summary>
        /// Propagated from <see cref="WorkbookEvaluator.clearAllCachedResultValues()" /> to clear locally cached data.
        /// Implementations must call the same method on all referenced <see cref="EvaluationSheet"/> instances, as well as clearing local caches.
        /// </summary>
        /// <see cref="WorkbookEvaluator.clearAllCachedResultValues()" />
        void ClearAllCachedResultValues();

        SpreadsheetVersion GetSpreadsheetVersion();
    }

    public class ExternalName
    {
        private String _nameName;
        private int _nameNumber;
        private int _ix;

        public ExternalName(String nameName, int nameNumber, int ix)
        {
            _nameName = nameName;
            _nameNumber = nameNumber;
            _ix = ix;
        }
        public String Name
        {
            get
            {
                return _nameName;
            }
        }
        public int Number
        {
            get
            {
                return _nameNumber;
            }
        }
        public int Ix
        {
            get
            {
                return _ix;
            }
        }
    }
}