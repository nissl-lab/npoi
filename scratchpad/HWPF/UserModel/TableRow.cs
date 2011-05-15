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


            _tprops = TableSprmUncompressor.uncompressTAP(_papx.ToByteArray(), 2);
            _levelNum = levelNum;
            _cells = new TableCell[_tprops.GetItcMac()];

            int start = 0;
            int end = 0;

            for (int cellIndex = 0; cellIndex < _cells.Length; cellIndex++)
            {
                Paragraph p = GetParagraph(start);
                String s = p.Text;

                while (!((s[s.Length - 1] == TABLE_CELL_MARK) ||
                          p.IsEmbeddedCellMark() && p.GetTableLevel() == levelNum))
                {
                    end++;
                    p = GetParagraph(end);
                    s = p.Text;
                }

                // Create it for the correct paragraph range
                _cells[cellIndex] = new TableCell(start, end, this, levelNum,
                                                  _tprops.GetRgtc()[cellIndex],
                                                  _tprops.GetRgdxaCenter()[cellIndex],
                                                  _tprops.GetRgdxaCenter()[cellIndex + 1] - _tprops.GetRgdxaCenter()[cellIndex]);
                // Now we've decided where everything is, tweak the
                //  record of the paragraph end so that the
                //  paragraph level counts work
                // This is a bit hacky, we really need a better fix...
                _cells[cellIndex]._parEnd++;

                // Next!
                end++;
                start = end;
            }
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

        public int numCells()
        {
            return _cells.Length;
        }

        public TableCell GetCell(int index)
        {
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