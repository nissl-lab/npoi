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

namespace NPOI.SS.Util
{


    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NUnit.Framework;

    /// <summary>
    /// Tests Spreadsheet PropertyTemplate
    /// </summary>
    /// <see cref="NPOI.SS.Util.PropertyTemplate" />
    [TestFixture]
    public sealed class TestPropertyTemplate
    {
        [Test]
        public void GetNumBorders()
        {

            CellRangeAddress a1 = new CellRangeAddress(0, 0, 0, 0);
            PropertyTemplate pt = new PropertyTemplate();
            pt.DrawBorders(a1, BorderStyle.Thin, BorderExtent.TOP);
            Assert.AreEqual(1, pt.GetNumBorders(0, 0));
            pt.DrawBorders(a1, BorderStyle.Medium, BorderExtent.BOTTOM);
            Assert.AreEqual(2, pt.GetNumBorders(0, 0));
            pt.DrawBorders(a1, BorderStyle.Medium, BorderExtent.NONE);
            Assert.AreEqual(0, pt.GetNumBorders(0, 0));
        }

        [Test]
        public void GetNumBorderColors()
        {

            CellRangeAddress a1 = new CellRangeAddress(0, 0, 0, 0);
            PropertyTemplate pt = new PropertyTemplate();
            pt.DrawBorderColors(a1, IndexedColors.Red.Index, BorderExtent.TOP);
            Assert.AreEqual(1, pt.GetNumBorderColors(0, 0));
            pt.DrawBorderColors(a1, IndexedColors.Red.Index, BorderExtent.BOTTOM);
            Assert.AreEqual(2, pt.GetNumBorderColors(0, 0));
            pt.DrawBorderColors(a1, IndexedColors.Red.Index, BorderExtent.NONE);
            Assert.AreEqual(0, pt.GetNumBorderColors(0, 0));
        }

        [Test]
        public void GetTemplateProperties()
        {

            CellRangeAddress a1 = new CellRangeAddress(0, 0, 0, 0);
            PropertyTemplate pt = new PropertyTemplate();
            pt.DrawBorders(a1, BorderStyle.Thin, BorderExtent.TOP);
            Assert.AreEqual(BorderStyle.Thin,
                    pt.GetBorderStyle(0, 0, CellUtil.BORDER_TOP));
            pt.DrawBorders(a1, BorderStyle.Medium, BorderExtent.BOTTOM);
            Assert.AreEqual(BorderStyle.Medium,
                    pt.GetBorderStyle(0, 0, CellUtil.BORDER_BOTTOM));
            pt.DrawBorderColors(a1, IndexedColors.Red.Index, BorderExtent.TOP);
            Assert.AreEqual(IndexedColors.Red.Index,
                    pt.GetTemplateProperty(0, 0, CellUtil.TOP_BORDER_COLOR));
            pt.DrawBorderColors(a1, IndexedColors.Blue.Index, BorderExtent.BOTTOM);
            Assert.AreEqual(IndexedColors.Blue.Index,
                    pt.GetTemplateProperty(0, 0, CellUtil.BOTTOM_BORDER_COLOR));
        }

        [Test]
        public void DrawBorders()
        {

            CellRangeAddress a1c3 = new CellRangeAddress(0, 2, 0, 2);
            PropertyTemplate pt = new PropertyTemplate();
            pt.DrawBorders(a1c3, BorderStyle.Thin,
                    BorderExtent.ALL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    Assert.AreEqual(4, pt.GetNumBorders(i, j));
                    Assert.AreEqual(BorderStyle.Thin,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_TOP));
                    Assert.AreEqual(BorderStyle.Thin,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_BOTTOM));
                    Assert.AreEqual(BorderStyle.Thin,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_LEFT));
                    Assert.AreEqual(BorderStyle.Thin,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_RIGHT));
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.Medium,
                    BorderExtent.OUTSIDE);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    Assert.AreEqual(4, pt.GetNumBorders(i, j));
                    if(i == 0)
                    {
                        if(j == 0)
                        {
                            Assert.AreEqual(BorderStyle.Medium,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_TOP));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_BOTTOM));
                            Assert.AreEqual(BorderStyle.Medium,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_LEFT));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_RIGHT));
                        }
                        else if(j == 2)
                        {
                            Assert.AreEqual(BorderStyle.Medium,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_TOP));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_BOTTOM));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_LEFT));
                            Assert.AreEqual(BorderStyle.Medium,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_RIGHT));
                        }
                        else
                        {
                            Assert.AreEqual(BorderStyle.Medium,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_TOP));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_BOTTOM));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_LEFT));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_RIGHT));
                        }
                    }
                    else if(i == 2)
                    {
                        if(j == 0)
                        {
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_TOP));
                            Assert.AreEqual(BorderStyle.Medium,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_BOTTOM));
                            Assert.AreEqual(BorderStyle.Medium,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_LEFT));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_RIGHT));
                        }
                        else if(j == 2)
                        {
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_TOP));
                            Assert.AreEqual(BorderStyle.Medium,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_BOTTOM));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_LEFT));
                            Assert.AreEqual(BorderStyle.Medium,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_RIGHT));
                        }
                        else
                        {
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_TOP));
                            Assert.AreEqual(BorderStyle.Medium,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_BOTTOM));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_LEFT));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_RIGHT));
                        }
                    }
                    else
                    {
                        if(j == 0)
                        {
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_TOP));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_BOTTOM));
                            Assert.AreEqual(BorderStyle.Medium,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_LEFT));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_RIGHT));
                        }
                        else if(j == 2)
                        {
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_TOP));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_BOTTOM));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_LEFT));
                            Assert.AreEqual(BorderStyle.Medium,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_RIGHT));
                        }
                        else
                        {
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_TOP));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_BOTTOM));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_LEFT));
                            Assert.AreEqual(BorderStyle.Thin,
                                    pt.GetBorderStyle(i, j,
                                            CellUtil.BORDER_RIGHT));
                        }
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    Assert.AreEqual(0, pt.GetNumBorders(i, j));
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.Medium,
                    BorderExtent.TOP);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(i == 0)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(BorderStyle.Medium,
                                pt.GetBorderStyle(i, j, CellUtil.BORDER_TOP));
                    }
                    else
                    {
                        Assert.AreEqual(0, pt.GetNumBorders(i, j));
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorders(a1c3, BorderStyle.Medium,
                    BorderExtent.BOTTOM);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(i == 2)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(BorderStyle.Medium, pt
                                .GetBorderStyle(i, j, CellUtil.BORDER_BOTTOM));
                    }
                    else
                    {
                        Assert.AreEqual(0, pt.GetNumBorders(i, j));
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorders(a1c3, BorderStyle.Medium,
                    BorderExtent.LEFT);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(j == 0)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(BorderStyle.Medium,
                                pt.GetBorderStyle(i, j, CellUtil.BORDER_LEFT));
                    }
                    else
                    {
                        Assert.AreEqual(0, pt.GetNumBorders(i, j));
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorders(a1c3, BorderStyle.Medium,
                    BorderExtent.RIGHT);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(j == 2)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(BorderStyle.Medium, pt
                                .GetBorderStyle(i, j, CellUtil.BORDER_RIGHT));
                    }
                    else
                    {
                        Assert.AreEqual(0, pt.GetNumBorders(i, j));
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorders(a1c3, BorderStyle.Medium,
                    BorderExtent.HORIZONTAL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    Assert.AreEqual(2, pt.GetNumBorders(i, j));
                    Assert.AreEqual(BorderStyle.Medium,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_TOP));
                    Assert.AreEqual(BorderStyle.Medium,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_BOTTOM));
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorders(a1c3, BorderStyle.Medium,
                    BorderExtent.INSIDE_HORIZONTAL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(i == 0)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(BorderStyle.Medium, pt
                                .GetBorderStyle(i, j, CellUtil.BORDER_BOTTOM));
                    }
                    else if(i == 2)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(BorderStyle.Medium,
                                pt.GetBorderStyle(i, j, CellUtil.BORDER_TOP));
                    }
                    else
                    {
                        Assert.AreEqual(2, pt.GetNumBorders(i, j));
                        Assert.AreEqual(BorderStyle.Medium,
                                pt.GetBorderStyle(i, j, CellUtil.BORDER_TOP));
                        Assert.AreEqual(BorderStyle.Medium, pt
                                .GetBorderStyle(i, j, CellUtil.BORDER_BOTTOM));
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorders(a1c3, BorderStyle.Medium,
                    BorderExtent.OUTSIDE_HORIZONTAL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(i == 0)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(BorderStyle.Medium,
                                pt.GetBorderStyle(i, j, CellUtil.BORDER_TOP));
                    }
                    else if(i == 2)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(BorderStyle.Medium, pt
                                .GetBorderStyle(i, j, CellUtil.BORDER_BOTTOM));
                    }
                    else
                    {
                        Assert.AreEqual(0, pt.GetNumBorders(i, j));
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorders(a1c3, BorderStyle.Medium,
                    BorderExtent.VERTICAL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    Assert.AreEqual(2, pt.GetNumBorders(i, j));
                    Assert.AreEqual(BorderStyle.Medium,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_LEFT));
                    Assert.AreEqual(BorderStyle.Medium,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_RIGHT));
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorders(a1c3, BorderStyle.Medium,
                    BorderExtent.INSIDE_VERTICAL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(j == 0)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(BorderStyle.Medium, pt
                                .GetBorderStyle(i, j, CellUtil.BORDER_RIGHT));
                    }
                    else if(j == 2)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(BorderStyle.Medium,
                                pt.GetBorderStyle(i, j, CellUtil.BORDER_LEFT));
                    }
                    else
                    {
                        Assert.AreEqual(2, pt.GetNumBorders(i, j));
                        Assert.AreEqual(BorderStyle.Medium,
                                pt.GetBorderStyle(i, j, CellUtil.BORDER_LEFT));
                        Assert.AreEqual(BorderStyle.Medium, pt
                                .GetBorderStyle(i, j, CellUtil.BORDER_RIGHT));
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorders(a1c3, BorderStyle.Medium,
                    BorderExtent.OUTSIDE_VERTICAL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(j == 0)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(BorderStyle.Medium,
                                pt.GetBorderStyle(i, j, CellUtil.BORDER_LEFT));
                    }
                    else if(j == 2)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(BorderStyle.Medium, pt
                                .GetBorderStyle(i, j, CellUtil.BORDER_RIGHT));
                    }
                    else
                    {
                        Assert.AreEqual(0, pt.GetNumBorders(i, j));
                    }
                }
            }
        }

        [Test]
        public void DrawBorderColors()
        {

            CellRangeAddress a1c3 = new CellRangeAddress(0, 2, 0, 2);
            PropertyTemplate pt = new PropertyTemplate();
            pt.DrawBorderColors(a1c3, IndexedColors.Red.Index,
                    BorderExtent.ALL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    Assert.AreEqual(4, pt.GetNumBorders(i, j));
                    Assert.AreEqual(4, pt.GetNumBorderColors(i, j));
                    Assert.AreEqual(IndexedColors.Red.Index, pt
                            .GetTemplateProperty(i, j, CellUtil.TOP_BORDER_COLOR));
                    Assert.AreEqual(IndexedColors.Red.Index,
                            pt.GetTemplateProperty(i, j,
                                    CellUtil.BOTTOM_BORDER_COLOR));
                    Assert.AreEqual(IndexedColors.Red.Index, pt
                            .GetTemplateProperty(i, j, CellUtil.LEFT_BORDER_COLOR));
                    Assert.AreEqual(IndexedColors.Red.Index,
                            pt.GetTemplateProperty(i, j,
                                    CellUtil.RIGHT_BORDER_COLOR));
                }
            }
            pt.DrawBorderColors(a1c3, IndexedColors.Blue.Index,
                    BorderExtent.OUTSIDE);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    Assert.AreEqual(4, pt.GetNumBorders(i, j));
                    Assert.AreEqual(4, pt.GetNumBorderColors(i, j));
                    if(i == 0)
                    {
                        if(j == 0)
                        {
                            Assert.AreEqual(IndexedColors.Blue.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.TOP_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.BOTTOM_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Blue.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.LEFT_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.RIGHT_BORDER_COLOR));
                        }
                        else if(j == 2)
                        {
                            Assert.AreEqual(IndexedColors.Blue.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.TOP_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.BOTTOM_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.LEFT_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Blue.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.RIGHT_BORDER_COLOR));
                        }
                        else
                        {
                            Assert.AreEqual(IndexedColors.Blue.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.TOP_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.BOTTOM_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.LEFT_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.RIGHT_BORDER_COLOR));
                        }
                    }
                    else if(i == 2)
                    {
                        if(j == 0)
                        {
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.TOP_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Blue.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.BOTTOM_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Blue.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.LEFT_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.RIGHT_BORDER_COLOR));
                        }
                        else if(j == 2)
                        {
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.TOP_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Blue.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.BOTTOM_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.LEFT_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Blue.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.RIGHT_BORDER_COLOR));
                        }
                        else
                        {
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.TOP_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Blue.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.BOTTOM_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.LEFT_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.RIGHT_BORDER_COLOR));
                        }
                    }
                    else
                    {
                        if(j == 0)
                        {
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.TOP_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.BOTTOM_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Blue.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.LEFT_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.RIGHT_BORDER_COLOR));
                        }
                        else if(j == 2)
                        {
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.TOP_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.BOTTOM_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.LEFT_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Blue.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.RIGHT_BORDER_COLOR));
                        }
                        else
                        {
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.TOP_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.BOTTOM_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.LEFT_BORDER_COLOR));
                            Assert.AreEqual(IndexedColors.Red.Index,
                                    pt.GetTemplateProperty(i, j,
                                            CellUtil.RIGHT_BORDER_COLOR));
                        }
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Automatic.Index,
                    BorderExtent.NONE);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    Assert.AreEqual(0, pt.GetNumBorders(i, j));
                    Assert.AreEqual(0, pt.GetNumBorderColors(i, j));
                }
            }
            pt.DrawBorderColors(a1c3, IndexedColors.Blue.Index,
                    BorderExtent.TOP);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(i == 0)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(1, pt.GetNumBorderColors(i, j));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.TOP_BORDER_COLOR));
                    }
                    else
                    {
                        Assert.AreEqual(0, pt.GetNumBorders(i, j));
                        Assert.AreEqual(0, pt.GetNumBorderColors(i, j));
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Automatic.Index,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Blue.Index,
                    BorderExtent.BOTTOM);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(i == 2)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(1, pt.GetNumBorderColors(i, j));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.BOTTOM_BORDER_COLOR));
                    }
                    else
                    {
                        Assert.AreEqual(0, pt.GetNumBorders(i, j));
                        Assert.AreEqual(0, pt.GetNumBorderColors(i, j));
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Automatic.Index,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Blue.Index,
                    BorderExtent.LEFT);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(j == 0)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(1, pt.GetNumBorderColors(i, j));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.LEFT_BORDER_COLOR));
                    }
                    else
                    {
                        Assert.AreEqual(0, pt.GetNumBorders(i, j));
                        Assert.AreEqual(0, pt.GetNumBorderColors(i, j));
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Automatic.Index,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Blue.Index,
                    BorderExtent.RIGHT);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(j == 2)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(1, pt.GetNumBorderColors(i, j));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.RIGHT_BORDER_COLOR));
                    }
                    else
                    {
                        Assert.AreEqual(0, pt.GetNumBorders(i, j));
                        Assert.AreEqual(0, pt.GetNumBorderColors(i, j));
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Automatic.Index,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Blue.Index,
                    BorderExtent.HORIZONTAL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    Assert.AreEqual(2, pt.GetNumBorders(i, j));
                    Assert.AreEqual(2, pt.GetNumBorderColors(i, j));
                    Assert.AreEqual(IndexedColors.Blue.Index, pt
                            .GetTemplateProperty(i, j, CellUtil.TOP_BORDER_COLOR));
                    Assert.AreEqual(IndexedColors.Blue.Index,
                            pt.GetTemplateProperty(i, j,
                                    CellUtil.BOTTOM_BORDER_COLOR));
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Automatic.Index,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Blue.Index,
                    BorderExtent.INSIDE_HORIZONTAL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(i == 0)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(1, pt.GetNumBorderColors(i, j));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.BOTTOM_BORDER_COLOR));
                    }
                    else if(i == 2)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(1, pt.GetNumBorderColors(i, j));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.TOP_BORDER_COLOR));
                    }
                    else
                    {
                        Assert.AreEqual(2, pt.GetNumBorders(i, j));
                        Assert.AreEqual(2, pt.GetNumBorderColors(i, j));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.TOP_BORDER_COLOR));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.BOTTOM_BORDER_COLOR));
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Automatic.Index,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Blue.Index,
                    BorderExtent.OUTSIDE_HORIZONTAL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(i == 0)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(1, pt.GetNumBorderColors(i, j));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.TOP_BORDER_COLOR));
                    }
                    else if(i == 2)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(1, pt.GetNumBorderColors(i, j));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.BOTTOM_BORDER_COLOR));
                    }
                    else
                    {
                        Assert.AreEqual(0, pt.GetNumBorders(i, j));
                        Assert.AreEqual(0, pt.GetNumBorderColors(i, j));
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Automatic.Index,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Blue.Index,
                    BorderExtent.VERTICAL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    Assert.AreEqual(2, pt.GetNumBorders(i, j));
                    Assert.AreEqual(2, pt.GetNumBorderColors(i, j));
                    Assert.AreEqual(IndexedColors.Blue.Index, pt
                            .GetTemplateProperty(i, j, CellUtil.LEFT_BORDER_COLOR));
                    Assert.AreEqual(IndexedColors.Blue.Index,
                            pt.GetTemplateProperty(i, j,
                                    CellUtil.RIGHT_BORDER_COLOR));
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Automatic.Index,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Blue.Index,
                    BorderExtent.INSIDE_VERTICAL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(j == 0)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(1, pt.GetNumBorderColors(i, j));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.RIGHT_BORDER_COLOR));
                    }
                    else if(j == 2)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(1, pt.GetNumBorderColors(i, j));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.LEFT_BORDER_COLOR));
                    }
                    else
                    {
                        Assert.AreEqual(2, pt.GetNumBorders(i, j));
                        Assert.AreEqual(2, pt.GetNumBorderColors(i, j));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.LEFT_BORDER_COLOR));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.RIGHT_BORDER_COLOR));
                    }
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Automatic.Index,
                    BorderExtent.NONE);
            pt.DrawBorderColors(a1c3, IndexedColors.Blue.Index,
                    BorderExtent.OUTSIDE_VERTICAL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    if(j == 0)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(1, pt.GetNumBorderColors(i, j));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.LEFT_BORDER_COLOR));
                    }
                    else if(j == 2)
                    {
                        Assert.AreEqual(1, pt.GetNumBorders(i, j));
                        Assert.AreEqual(1, pt.GetNumBorderColors(i, j));
                        Assert.AreEqual(IndexedColors.Blue.Index,
                                pt.GetTemplateProperty(i, j,
                                        CellUtil.RIGHT_BORDER_COLOR));
                    }
                    else
                    {
                        Assert.AreEqual(0, pt.GetNumBorders(i, j));
                        Assert.AreEqual(0, pt.GetNumBorderColors(i, j));
                    }
                }
            }
        }

        [Test]
        public void DrawBordersWithColors()
        {

            CellRangeAddress a1c3 = new CellRangeAddress(0, 2, 0, 2);
            PropertyTemplate pt = new PropertyTemplate();

            pt.DrawBorders(a1c3, BorderStyle.Medium, IndexedColors.Red.Index, BorderExtent.ALL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    Assert.AreEqual(4, pt.GetNumBorders(i, j));
                    Assert.AreEqual(4, pt.GetNumBorderColors(i, j));
                    Assert.AreEqual(BorderStyle.Medium,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_TOP));
                    Assert.AreEqual(BorderStyle.Medium,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_BOTTOM));
                    Assert.AreEqual(BorderStyle.Medium,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_LEFT));
                    Assert.AreEqual(BorderStyle.Medium,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_RIGHT));
                    Assert.AreEqual(IndexedColors.Red.Index, pt
                            .GetTemplateProperty(i, j, CellUtil.TOP_BORDER_COLOR));
                    Assert.AreEqual(IndexedColors.Red.Index,
                            pt.GetTemplateProperty(i, j,
                                    CellUtil.BOTTOM_BORDER_COLOR));
                    Assert.AreEqual(IndexedColors.Red.Index, pt
                            .GetTemplateProperty(i, j, CellUtil.LEFT_BORDER_COLOR));
                    Assert.AreEqual(IndexedColors.Red.Index,
                            pt.GetTemplateProperty(i, j,
                                    CellUtil.RIGHT_BORDER_COLOR));
                }
            }
            pt.DrawBorders(a1c3, BorderStyle.None, BorderExtent.NONE);
            pt.DrawBorders(a1c3, BorderStyle.None, IndexedColors.Red.Index, BorderExtent.ALL);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    Assert.AreEqual(4, pt.GetNumBorders(i, j));
                    Assert.AreEqual(0, pt.GetNumBorderColors(i, j));
                    Assert.AreEqual(BorderStyle.None,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_TOP));
                    Assert.AreEqual(BorderStyle.None,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_BOTTOM));
                    Assert.AreEqual(BorderStyle.None,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_LEFT));
                    Assert.AreEqual(BorderStyle.None,
                            pt.GetBorderStyle(i, j, CellUtil.BORDER_RIGHT));
                }
            }
        }

        [Test]
        public void ApplyBorders()
        {

            CellRangeAddress a1c3 = new CellRangeAddress(0, 2, 0, 2);
            CellRangeAddress b2 = new CellRangeAddress(1, 1, 1, 1);
            PropertyTemplate pt = new PropertyTemplate();
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();

            pt.DrawBorders(a1c3, BorderStyle.Thin, IndexedColors.Red.Index, BorderExtent.ALL);
            pt.applyBorders(sheet);

            foreach(IRow row in sheet)
            {
                foreach(ICell cell in row)
                {
                    ICellStyle cs = cell.CellStyle;
                    Assert.AreEqual(BorderStyle.Thin, cs.BorderTop);
                    Assert.AreEqual(IndexedColors.Red.Index, cs.TopBorderColor);
                    Assert.AreEqual(BorderStyle.Thin, cs.BorderBottom);
                    Assert.AreEqual(IndexedColors.Red.Index, cs.BottomBorderColor);
                    Assert.AreEqual(BorderStyle.Thin, cs.BorderLeft);
                    Assert.AreEqual(IndexedColors.Red.Index, cs.LeftBorderColor);
                    Assert.AreEqual(BorderStyle.Thin, cs.BorderRight);
                    Assert.AreEqual(IndexedColors.Red.Index, cs.RightBorderColor);
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
                        Assert.AreEqual(BorderStyle.Thin, cs.BorderTop);
                        Assert.AreEqual(IndexedColors.Red.Index, cs.TopBorderColor);
                    }
                    else
                    {
                        Assert.AreEqual(BorderStyle.None, cs.BorderTop);
                    }
                    if(cell.ColumnIndex != 1 || row.RowNum == 2)
                    {
                        Assert.AreEqual(BorderStyle.Thin, cs.BorderBottom);
                        Assert.AreEqual(IndexedColors.Red.Index, cs.BottomBorderColor);
                    }
                    else
                    {
                        Assert.AreEqual(BorderStyle.None, cs.BorderBottom);
                    }
                    if(cell.ColumnIndex == 0 || row.RowNum != 1)
                    {
                        Assert.AreEqual(BorderStyle.Thin, cs.BorderLeft);
                        Assert.AreEqual(IndexedColors.Red.Index, cs.LeftBorderColor);
                    }
                    else
                    {
                        Assert.AreEqual(BorderStyle.None, cs.BorderLeft);
                    }
                    if(cell.ColumnIndex == 2 || row.RowNum != 1)
                    {
                        Assert.AreEqual(BorderStyle.Thin, cs.BorderRight);
                        Assert.AreEqual(IndexedColors.Red.Index, cs.RightBorderColor);
                    }
                    else
                    {
                        Assert.AreEqual(BorderStyle.None, cs.BorderRight);
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
            Assert.AreNotSame(pt2, pt);
            for(int i = 0; i <= 2; i++)
            {
                for(int j = 0; j <= 2; j++)
                {
                    Assert.AreEqual(4, pt2.GetNumBorderColors(i, j));
                    Assert.AreEqual(4, pt2.GetNumBorderColors(i, j));
                }
            }

            CellRangeAddress b2 = new CellRangeAddress(1,1,1,1);
            pt2.DrawBorders(b2, BorderStyle.Thin, BorderExtent.ALL);

            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            pt.applyBorders(sheet);

            foreach(IRow row in sheet)
            {
                foreach(ICell cell in row)
                {
                    ICellStyle cs = cell.CellStyle;
                    Assert.AreEqual(BorderStyle.Medium, cs.BorderTop);
                    Assert.AreEqual(BorderStyle.Medium, cs.BorderBottom);
                    Assert.AreEqual(BorderStyle.Medium, cs.BorderLeft);
                    Assert.AreEqual(BorderStyle.Medium, cs.BorderRight);
                    Assert.AreEqual(IndexedColors.Red.Index, cs.TopBorderColor);
                    Assert.AreEqual(IndexedColors.Red.Index, cs.BottomBorderColor);
                    Assert.AreEqual(IndexedColors.Red.Index, cs.LeftBorderColor);
                    Assert.AreEqual(IndexedColors.Red.Index, cs.RightBorderColor);
                }
            }

            wb.Close();
        }
    }
}

