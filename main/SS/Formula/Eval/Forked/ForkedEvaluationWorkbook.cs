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


namespace NPOI.SS.Formula.Eval.Forked
{
    using System;
    using System.Collections.Generic;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Formula.Udf;
    using NPOI.SS.UserModel;

    /**
     * Represents a workbook being used for forked Evaluation. Most operations are delegated to the
     * shared master workbook, except those that potentially involve cell values that may have been
     * updated After a call to {@link #getOrCreateUpdatableCell(String, int, int)}.
     *
     * @author Josh Micich
     */
    class ForkedEvaluationWorkbook : IEvaluationWorkbook
    {

        private IEvaluationWorkbook _masterBook;
        private Dictionary<String, ForkedEvaluationSheet> _sharedSheetsByName;

        public ForkedEvaluationWorkbook(IEvaluationWorkbook master)
        {
            _masterBook = master;
            _sharedSheetsByName = new Dictionary<String, ForkedEvaluationSheet>();
        }

        public ForkedEvaluationCell GetOrCreateUpdatableCell(String sheetName, int rowIndex,
                int columnIndex)
        {
            ForkedEvaluationSheet sheet = GetSharedSheet(sheetName);
            return sheet.GetOrCreateUpdatableCell(rowIndex, columnIndex);
        }

        public IEvaluationCell GetEvaluationCell(String sheetName, int rowIndex, int columnIndex)
        {
            ForkedEvaluationSheet sheet = GetSharedSheet(sheetName);
            return sheet.GetCell(rowIndex, columnIndex);
        }

        private ForkedEvaluationSheet GetSharedSheet(String sheetName)
        {
            ForkedEvaluationSheet result = null;
            if(_sharedSheetsByName.ContainsKey(sheetName))
                result = _sharedSheetsByName[(sheetName)];
            if (result == null)
            {
                result = new ForkedEvaluationSheet(_masterBook.GetSheet(_masterBook
                        .GetSheetIndex(sheetName)));
                if (_sharedSheetsByName.ContainsKey(sheetName))
                {
                    _sharedSheetsByName[sheetName] = result;
                }
                else
                {
                    _sharedSheetsByName.Add(sheetName, result);
                }
            }
            return result;
        }

        public void CopyUpdatedCells(IWorkbook workbook)
        {
            String[] sheetNames = new String[_sharedSheetsByName.Count];
            _sharedSheetsByName.Keys.CopyTo(sheetNames, 0);
            OrderedSheet[] oss = new OrderedSheet[sheetNames.Length];
            for (int i = 0; i < sheetNames.Length; i++)
            {
                String sheetName = sheetNames[i];
                oss[i] = new OrderedSheet(sheetName, _masterBook.GetSheetIndex(sheetName));
            }
            for (int i = 0; i < oss.Length; i++)
            {
                String sheetName = oss[i].GetSheetName();
                ForkedEvaluationSheet sheet = _sharedSheetsByName[(sheetName)];
                sheet.CopyUpdatedCells(workbook.GetSheet(sheetName));
            }
        }

        public int ConvertFromExternSheetIndex(int externSheetIndex)
        {
            return _masterBook.ConvertFromExternSheetIndex(externSheetIndex);
        }

        public ExternalSheet GetExternalSheet(int externSheetIndex)
        {
            return _masterBook.GetExternalSheet(externSheetIndex);
        }
        public ExternalSheet GetExternalSheet(String firstSheetName, string lastSheetName, int externalWorkbookNumber)
        {
            return _masterBook.GetExternalSheet(firstSheetName, lastSheetName, externalWorkbookNumber);
        }
        public Ptg[] GetFormulaTokens(IEvaluationCell cell)
        {
            if (cell is ForkedEvaluationCell)
            {
                // doesn't happen yet because formulas cannot be modified from the master workbook
                throw new Exception("Updated formulas not supported yet");
            }
            return _masterBook.GetFormulaTokens(cell);
        }

        public IEvaluationName GetName(NamePtg namePtg)
        {
            return _masterBook.GetName(namePtg);
        }

        public IEvaluationName GetName(String name, int sheetIndex)
        {
            return _masterBook.GetName(name, sheetIndex);
        }

        public IEvaluationSheet GetSheet(int sheetIndex)
        {
            return GetSharedSheet(GetSheetName(sheetIndex));
        }

        public ExternalName GetExternalName(int externSheetIndex, int externNameIndex)
        {
            return _masterBook.GetExternalName(externSheetIndex, externNameIndex);
        }
        public ExternalName GetExternalName(String nameName, String sheetName, int externalWorkbookNumber)
        {
            return _masterBook.GetExternalName(nameName, sheetName, externalWorkbookNumber);
        }
        public int GetSheetIndex(IEvaluationSheet sheet)
        {
            if (sheet is ForkedEvaluationSheet)
            {
                ForkedEvaluationSheet mes = (ForkedEvaluationSheet)sheet;
                return mes.GetSheetIndex(_masterBook);
            }
            return _masterBook.GetSheetIndex(sheet);
        }

        public int GetSheetIndex(String sheetName)
        {
            return _masterBook.GetSheetIndex(sheetName);
        }

        public String GetSheetName(int sheetIndex)
        {
            return _masterBook.GetSheetName(sheetIndex);
        }

        public String ResolveNameXText(NameXPtg ptg)
        {
            return _masterBook.ResolveNameXText(ptg);
        }

        public UDFFinder GetUDFFinder()
        {
            return _masterBook.GetUDFFinder();
        }

        private class OrderedSheet : IComparable<OrderedSheet>
        {
            private String _sheetName;
            private int _index;

            public OrderedSheet(String sheetName, int index)
            {
                _sheetName = sheetName;
                _index = index;
            }
            public String GetSheetName()
            {
                return _sheetName;
            }
            public int CompareTo(OrderedSheet o)
            {
                return _index - o._index;
            }
        }

    }

}