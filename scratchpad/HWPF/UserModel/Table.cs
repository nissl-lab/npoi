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

using System.Collections;
using System.Collections.Generic;
namespace NPOI.HWPF.UserModel
{

    public class Table : Range
    {
        List<TableRow> _rows;

        internal Table(int startIdx, int endIdx, Range parent, int levelNum)
            : base(startIdx, endIdx,  parent)
        {
              _tableLevel = levelNum;
              InitRows();
        }
        private int _tableLevel;
        private bool _rowsFound = false;
        private void InitRows()
        {
            if (_rowsFound)
                return;

            _rows = new List<TableRow>();
            int rowStart = 0;
            int rowEnd = 0;

            int numParagraphs = NumParagraphs;
            while (rowEnd < numParagraphs)
            {
                Paragraph startRowP = GetParagraph(rowStart);
                Paragraph endRowP = GetParagraph(rowEnd);
                rowEnd++;
                if (endRowP.IsTableRowEnd()
                        && endRowP.GetTableLevel() == _tableLevel)
                {
                    _rows.Add(new TableRow(startRowP.StartOffset, endRowP
                            .EndOffset, this, _tableLevel));
                    rowStart = rowEnd;
                }
            }
            _rowsFound = true;
        }
        public int NumRows
        {
            get
            {
                InitRows();
                return _rows.Count;
            }
        }

        public override int Type
        {
            get
            {
                return TYPE_TABLE;
            }
        }

        public TableRow GetRow(int index)
        {
            InitRows();
            return (TableRow)_rows[index];
        }

        public int TableLevel
        {
            get
            {
                return _tableLevel;
            }
        }
        protected override void Reset()
        {
            _rowsFound = false;
        }
       
    }

}