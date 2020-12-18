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

using System;
using NPOI.HWPF.SPRM;
using System.Collections.Generic;
namespace NPOI.HWPF.UserModel
{

    public class TableRow
      : Paragraph
    {
        private static char TABLE_CELL_MARK = '\u0007';

        private static short SPRM_TJC = 0x5400;
        private static short SPRM_DXAGAPHALF = unchecked((short)0x9602);
        private static short SPRM_FCANTSPLIT = 0x3403;
        private static short SPRM_FTABLEHEADER = 0x3404;
        private static short SPRM_DYAROWHEIGHT = unchecked((short)0x9407);

        int _levelNum;
        private TableProperties _tprops;
        private TableCell[] _cells;

        public TableRow(int startIdx, int endIdx, Table parent, int levelNum)
            : base(startIdx, endIdx, parent)
        {

            Paragraph last = GetParagraph(NumParagraphs - 1);
            _papx = last._papx;
            _tprops = TableSprmUncompressor.UncompressTAP(_papx);
            _levelNum = levelNum;
            initCells();

        }

            private void initCells()
    {
        if ( _cellsFound )
            return;

        short expectedCellsCount = _tprops.GetItcMac();

        int lastCellStart = 0;
        List<TableCell> cells = new List<TableCell>(
                expectedCellsCount + 1 );
        for ( int p = 0; p < NumParagraphs; p++ )
        {
            Paragraph paragraph = GetParagraph( p );
            String s = paragraph.Text;

            if ( ( ( s.Length > 0 && s[s.Length - 1]== TABLE_CELL_MARK ) || paragraph
                    .IsEmbeddedCellMark() )
                    && paragraph.GetTableLevel() == _levelNum )
            {
                TableCellDescriptor tableCellDescriptor = _tprops.GetRgtc() != null
                        && _tprops.GetRgtc().Length > cells.Count ? _tprops
                        .GetRgtc()[cells.Count] : new TableCellDescriptor();
                short leftEdge = (_tprops.GetRgdxaCenter() != null
                        && _tprops.GetRgdxaCenter().Length > cells.Count) ? (short)_tprops
                        .GetRgdxaCenter()[cells.Count] : (short)0;
                short rightEdge = (_tprops.GetRgdxaCenter() != null
                        && _tprops.GetRgdxaCenter().Length > cells.Count + 1) ? (short)_tprops
                        .GetRgdxaCenter()[cells.Count + 1] : (short)0;

                TableCell tableCell = new TableCell( GetParagraph(
                        lastCellStart ).StartOffset, GetParagraph( p )
                        .EndOffset, this, _levelNum, tableCellDescriptor,
                        leftEdge, rightEdge - leftEdge );
                cells.Add( tableCell );
                lastCellStart = p + 1;
            }
        }

        if ( lastCellStart < ( NumParagraphs - 1 ) )
        {
            TableCellDescriptor tableCellDescriptor = _tprops.GetRgtc() != null
                    && _tprops.GetRgtc().Length > cells.Count ? _tprops
                    .GetRgtc()[cells.Count] : new TableCellDescriptor();
            short leftEdge = _tprops.GetRgdxaCenter() != null
                    && _tprops.GetRgdxaCenter().Length > cells.Count ? (short)_tprops
                    .GetRgdxaCenter()[cells.Count] : (short)0;
            short rightEdge = _tprops.GetRgdxaCenter() != null
                    && _tprops.GetRgdxaCenter().Length > cells.Count + 1 ? (short)_tprops
                    .GetRgdxaCenter()[cells.Count + 1] : (short)0;

            TableCell tableCell = new TableCell( lastCellStart,
                    ( NumParagraphs - 1 ), this, _levelNum,
                    tableCellDescriptor, leftEdge, rightEdge - leftEdge );
            cells.Add( tableCell );
        }

        if ( cells.Count>0 )
        {
            TableCell lastCell = cells[cells.Count - 1];
            if ( lastCell.NumParagraphs == 1
                    && ( lastCell.GetParagraph( 0 ).IsTableRowEnd() ) )
            {
                // remove "fake" cell
                cells.RemoveAt( cells.Count - 1 );
            }
        }

        if ( cells.Count != expectedCellsCount )
        {
            _tprops.SetItcMac( (short) cells.Count);
        }

        _cells = cells.ToArray();
        _cellsFound = true;
    }
            private bool _cellsFound = false;
            protected void Reset()
            {
                _cellsFound = false;
            }
        public int GetRowJustification()
        {
            return _tprops.GetJc();
        }

        public void SetRowJustification(int jc)
        {
            _tprops.SetJc(jc);
            _papx.UpdateSprm(SPRM_TJC, (short)jc);
        }

        public int GetGapHalf()
        {
            return _tprops.GetDxaGapHalf();
        }

        public void SetGapHalf(int dxaGapHalf)
        {
            _tprops.SetDxaGapHalf(dxaGapHalf);
            _papx.UpdateSprm(SPRM_DXAGAPHALF, (short)dxaGapHalf);
        }

        public int GetRowHeight()
        {
            return _tprops.GetDyaRowHeight();
        }

        public void SetRowHeight(int dyaRowHeight)
        {
            _tprops.SetDyaRowHeight(dyaRowHeight);
            _papx.UpdateSprm(SPRM_DYAROWHEIGHT, (short)dyaRowHeight);
        }

        public bool cantSplit()
        {
            return _tprops.GetFCantSplit();
        }

        public void SetCantSplit(bool cantSplit)
        {
            _tprops.SetFCantSplit(cantSplit);
            _papx.UpdateSprm(SPRM_FCANTSPLIT, (byte)(cantSplit ? 1 : 0));
        }

        public bool isTableHeader()
        {
            return _tprops.GetFTableHeader();
        }

        public void SetTableHeader(bool tableHeader)
        {
            _tprops.SetFTableHeader(tableHeader);
            _papx.UpdateSprm(SPRM_FTABLEHEADER, (byte)(tableHeader ? 1 : 0));
        }

        public int NumCells()
        {
            initCells();
            return _cells.Length;
        }

        public TableCell GetCell(int index)
        {
            initCells();
            return _cells[index];
        }

        public override BorderCode GetTopBorder()
        {
            return _tprops.GetBrcBottom();
        }

        public override BorderCode GetBottomBorder()
        {
            return _tprops.GetBrcBottom();
        }

        public override BorderCode GetLeftBorder()
        {
            return _tprops.GetBrcLeft();
        }

        public override BorderCode GetRightBorder()
        {
            return _tprops.GetBrcRight();
        }

        public BorderCode GetHorizontalBorder()
        {
            return _tprops.GetBrcHorizontal();
        }

        public BorderCode GetVerticalBorder()
        {
            return _tprops.GetBrcVertical();
        }

        public override BorderCode GetBarBorder()
        {
            throw new NotImplementedException("not applicable for TableRow");
        }

    }

}