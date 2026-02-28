/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.SS.Util
{
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    [TestFixture]
    public class TestXSSFPropertyTemplate
    {

        [Test]
        public void ApplyBorders()
        {

            CellRangeAddress a1c3 = new CellRangeAddress(0, 2, 0, 2);
            CellRangeAddress b2 = new CellRangeAddress(1, 1, 1, 1);
            PropertyTemplate pt = new PropertyTemplate();
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();

            pt.DrawBorders(a1c3, BorderStyle.Thin, IndexedColors.Red.Index, BorderExtent.ALL);
            pt.applyBorders(sheet);

            foreach(IRow row in sheet)
            {
                foreach(ICell cell in row)
                {
                    ICellStyle cs = cell.CellStyle;
                    ClassicAssert.AreEqual(BorderStyle.Thin, cs.BorderTop);
                    ClassicAssert.AreEqual(IndexedColors.Red.Index, cs.TopBorderColor);
                    ClassicAssert.AreEqual(BorderStyle.Thin, cs.BorderBottom);
                    ClassicAssert.AreEqual(IndexedColors.Red.Index, cs.BottomBorderColor);
                    ClassicAssert.AreEqual(BorderStyle.Thin, cs.BorderLeft);
                    ClassicAssert.AreEqual(IndexedColors.Red.Index, cs.LeftBorderColor);
                    ClassicAssert.AreEqual(BorderStyle.Thin, cs.BorderRight);
                    ClassicAssert.AreEqual(IndexedColors.Red.Index, cs.RightBorderColor);
                }
            }

            pt.DrawBorders(b2, BorderStyle.None, BorderExtent.ALL);
            pt.applyBorders(sheet);

            foreach(IRow row in sheet)
            {
                foreach(ICell cell in row)
                {
                    ICellStyle cs = cell.CellStyle;
                    if(cell.ColumnIndex != 1 || row.RowNum == 0)
                    {
                        ClassicAssert.AreEqual(BorderStyle.Thin, cs.BorderTop);
                        ClassicAssert.AreEqual(IndexedColors.Red.Index, cs.TopBorderColor);
                    }
                    else
                    {
                        ClassicAssert.AreEqual(BorderStyle.None, cs.BorderTop);
                    }
                    if(cell.ColumnIndex != 1 || row.RowNum == 2)
                    {
                        ClassicAssert.AreEqual(BorderStyle.Thin, cs.BorderBottom);
                        ClassicAssert.AreEqual(IndexedColors.Red.Index, cs.BottomBorderColor);
                    }
                    else
                    {
                        ClassicAssert.AreEqual(BorderStyle.None, cs.BorderBottom);
                    }
                    if(cell.ColumnIndex == 0 || row.RowNum != 1)
                    {
                        ClassicAssert.AreEqual(BorderStyle.Thin, cs.BorderLeft);
                        ClassicAssert.AreEqual(IndexedColors.Red.Index, cs.LeftBorderColor);
                    }
                    else
                    {
                        ClassicAssert.AreEqual(BorderStyle.None, cs.BorderLeft);
                    }
                    if(cell.ColumnIndex == 2 || row.RowNum != 1)
                    {
                        ClassicAssert.AreEqual(BorderStyle.Thin, cs.BorderRight);
                        ClassicAssert.AreEqual(IndexedColors.Red.Index, cs.RightBorderColor);
                    }
                    else
                    {
                        ClassicAssert.AreEqual(BorderStyle.None, cs.BorderRight);
                    }
                }
            }

            wb.Close();
        }

        [Test]
        public void ClonePropertyTemplate()
        {

            CellRangeAddress a1c3 = new CellRangeAddress(0, 2, 0, 2);
            PropertyTemplate pt = new PropertyTemplate();
            pt.DrawBorders(a1c3, BorderStyle.Medium, IndexedColors.Red.Index, BorderExtent.ALL);
            PropertyTemplate pt2 = new PropertyTemplate(pt);
            ClassicAssert.AreNotSame(pt2, pt);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    ClassicAssert.AreEqual(4, pt2.GetNumBorderColors(i, j));
                    ClassicAssert.AreEqual(4, pt2.GetNumBorderColors(i, j));
                }
            }

            CellRangeAddress b2 = new CellRangeAddress(1,1,1,1);
            pt2.DrawBorders(b2, BorderStyle.Thin, BorderExtent.ALL);

            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            pt.applyBorders(sheet);

            foreach(IRow row in sheet)
            {
                foreach(ICell cell in row)
                {
                    ICellStyle cs = cell.CellStyle;
                    ClassicAssert.AreEqual(BorderStyle.Medium, cs.BorderTop);
                    ClassicAssert.AreEqual(BorderStyle.Medium, cs.BorderBottom);
                    ClassicAssert.AreEqual(BorderStyle.Medium, cs.BorderLeft);
                    ClassicAssert.AreEqual(BorderStyle.Medium, cs.BorderRight);
                    ClassicAssert.AreEqual(IndexedColors.Red.Index, cs.TopBorderColor);
                    ClassicAssert.AreEqual(IndexedColors.Red.Index, cs.BottomBorderColor);
                    ClassicAssert.AreEqual(IndexedColors.Red.Index, cs.LeftBorderColor);
                    ClassicAssert.AreEqual(IndexedColors.Red.Index, cs.RightBorderColor);
                }
            }

            wb.Close();
        }
    }
}

