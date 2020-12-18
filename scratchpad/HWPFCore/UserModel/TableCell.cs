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

    public class TableCell
      : Range
    {
        private int _levelNum;
        private TableCellDescriptor _tcd;
        private int _leftEdge;
        private int _width;

        public TableCell(int startIdx, int endIdx, TableRow parent, int levelNum, TableCellDescriptor tcd, int leftEdge, int width)
            : base(startIdx, endIdx,parent)
        {
            _tcd = tcd;
            _leftEdge = leftEdge;
            _width = width;
            _levelNum = levelNum;
        }

        public bool IsFirstMerged()
        {
            return _tcd.IsFFirstMerged();
        }

        public bool IsMerged()
        {
            return _tcd.IsFMerged();
        }

        public bool IsVertical()
        {
            return _tcd.IsFVertical();
        }

        public bool IsBackward()
        {
            return _tcd.IsFBackward();
        }

        public bool IsRotateFont()
        {
            return _tcd.IsFRotateFont();
        }

        public bool IsVerticallyMerged()
        {
            return _tcd.IsFVertMerge();
        }

        public bool IsFirstVerticallyMerged()
        {
            return _tcd.IsFVertRestart();
        }

        public byte GetVertAlign()
        {
            return _tcd.GetVertAlign();
        }

        public BorderCode GetBrcTop()
        {
            return _tcd.GetBrcTop();
        }

        public BorderCode GetBrcBottom()
        {
            return _tcd.GetBrcBottom();
        }

        public BorderCode GetBrcLeft()
        {
            return _tcd.GetBrcLeft();
        }

        public BorderCode GetBrcRight()
        {
            return _tcd.GetBrcRight();
        }

        public int GetLeftEdge() // twips
        {
            return _leftEdge;
        }

        public int GetWidth() // twips
        {
            return _width;
        }

        /** Returns the TableCellDescriptor for this cell.*/
        public TableCellDescriptor GetDescriptor()
        {
            return _tcd;
        }

    }

}