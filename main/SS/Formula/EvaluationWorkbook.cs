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
    using NPOI.SS.Formula.Udf;
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

    /**
     * Abstracts a workbook for the purpose of formula evaluation.<br/>
     * 
     * For POI internal use only
     * 
     * @author Josh Micich
     */
    public interface IEvaluationWorkbook
    {
        String GetSheetName(int sheetIndex);
        /**
         * @return -1 if the specified sheet is from a different book
         */
        int GetSheetIndex(IEvaluationSheet sheet);
        int GetSheetIndex(String sheetName);

        IEvaluationSheet GetSheet(int sheetIndex);

        /**
         * HSSF Only - fetch the external-style sheet details
         * <p>Return will have no workbook set if it's actually in our own workbook</p>
         */
        ExternalSheet GetExternalSheet(int externSheetIndex);
        /**
         * XSSF Only - fetch the external-style sheet details
         * <p>Return will have no workbook set if it's actually in our own workbook</p>
         */
        ExternalSheet GetExternalSheet(String firstSheetName, string lastSheetName, int externalWorkbookNumber);
        /**
         * HSSF Only - convert an external sheet index to an internal sheet index,
         *  for an external-style reference to one of this workbook's own sheets 
         */
        int ConvertFromExternSheetIndex(int externSheetIndex);
        /**
         * HSSF Only - fetch the external-style name details
         */
        ExternalName GetExternalName(int externSheetIndex, int externNameIndex);
        /**
         * XSSF Only - fetch the external-style name details
         */
        ExternalName GetExternalName(String nameName, String sheetName, int externalWorkbookNumber);
    
        IEvaluationName GetName(NamePtg namePtg);
        IEvaluationName GetName(String name, int sheetIndex);
        String ResolveNameXText(NameXPtg ptg);
        Ptg[] GetFormulaTokens(IEvaluationCell cell);
        UDFFinder GetUDFFinder();
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