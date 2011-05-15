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
namespace NPOI.HWPF.UserModel
{

    public class Table : Range
    {
        ArrayList _rows;

        internal Table(int startIdx, int endIdx, Range parent, int levelNum)
            : base(startIdx, endIdx, Range.TYPE_PARAGRAPH, parent)
        {

            _rows = new ArrayList();
            int numParagraphs = NumParagraphs;

            int rowStart = 0;
            int rowEnd = 0;

            while (rowEnd < numParagraphs)
            {
                Paragraph p = GetParagraph(rowEnd);
                rowEnd++;
                if (p.IsTableRowEnd() && p.GetTableLevel() == levelNum)
                {
                    _rows.Add(new TableRow(rowStart, rowEnd, this, levelNum));
                    rowStart = rowEnd;
                }
            }
        }

        public int NumRows
        {
            get
            {
                return _rows.Count;
            }
        }

        public int Type
        {
            get
            {
                return TYPE_TABLE;
            }
        }

        public TableRow GetRow(int index)
        {
            return (TableRow)_rows[index];
        }
    }

}