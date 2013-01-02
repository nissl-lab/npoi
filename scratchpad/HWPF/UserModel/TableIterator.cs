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

namespace NPOI.HWPF.UserModel
{

    public class TableIterator
    {
        Range _range;
        int _index;
        int _levelNum;

        TableIterator(Range range, int levelNum)
        {
            _range = range;
            _index = 0;
            _levelNum = levelNum;
        }

        public TableIterator(Range range)
            : this(range, 1)
        {

        }


        public bool HasNext()
        {
            int numParagraphs = _range.NumParagraphs;
            for (; _index < numParagraphs; _index++)
            {
                Paragraph paragraph = _range.GetParagraph(_index);
                if (paragraph.IsInTable() && paragraph.GetTableLevel() == _levelNum)
                {
                    return true;
                }
            }
            return false;
        }

        public Table Next()
        {
            int numParagraphs = _range.NumParagraphs;
            int startIndex = _index;
            int endIndex = _index;

            for (; _index < numParagraphs; _index++)
            {
                Paragraph paragraph = _range.GetParagraph(_index);
                if (!paragraph.IsInTable() || paragraph.GetTableLevel() < _levelNum)
                {
                    endIndex = _index;
                    break;
                }
            }
            return new Table(startIndex, endIndex, _range, _levelNum);
        }

    }
}
