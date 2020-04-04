/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.UserModel
{
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.CF;
    using NPOI.HSSF.Util;
    using NPOI.SS.UserModel;

    /**
     * High level representation for Border Formatting component
     * of Conditional Formatting Settings
     * 
     * @author Dmitriy Kumshayev
     *
     */
    public class HSSFBorderFormatting : IBorderFormatting
    {
        private HSSFWorkbook workbook;
        private CFRuleBase cfRuleRecord;
        private BorderFormatting borderFormatting;

        public HSSFBorderFormatting(CFRuleBase cfRuleRecord, HSSFWorkbook workbook)
        {
            this.workbook = workbook;
            this.cfRuleRecord = cfRuleRecord;
            this.borderFormatting = cfRuleRecord.BorderFormatting;
        }

        public BorderFormatting GetBorderFormattingBlock()
        {
            return borderFormatting;
        }

        public BorderStyle BorderBottom
        {
            get { return borderFormatting.BorderBottom; }
            set
            {
                borderFormatting.BorderBottom = value;
                if (value != BorderStyle.None)
                    cfRuleRecord.IsBottomBorderModified = true;
                else
                    cfRuleRecord.IsBottomBorderModified = (false);
            }
        }

        public BorderStyle BorderDiagonal
        {
            get { return borderFormatting.BorderDiagonal; }
            set {
                borderFormatting.BorderDiagonal = value;
                if (value != BorderStyle.None) {
                    cfRuleRecord.IsBottomLeftTopRightBorderModified = true;
                    cfRuleRecord.IsTopLeftBottomRightBorderModified = true;
                }
                else
                {
                    cfRuleRecord.IsBottomLeftTopRightBorderModified = (false);
                    cfRuleRecord.IsTopLeftBottomRightBorderModified = (false);
                }

            }
        }

        public BorderStyle BorderLeft
        {
            get { return borderFormatting.BorderLeft; }
            set {
                borderFormatting.BorderLeft = value;
                if (value != BorderStyle.None)
                    cfRuleRecord.IsLeftBorderModified = true;
                else
                {
                    cfRuleRecord.IsLeftBorderModified = (false);
                }
            }
        }

        public BorderStyle BorderRight
        {
            get { return borderFormatting.BorderRight; }
            set {
                borderFormatting.BorderRight = value;
                if (value != BorderStyle.None)
                    cfRuleRecord.IsRightBorderModified = true;
                else
                {
                    cfRuleRecord.IsRightBorderModified = (false);
                }
            }
        }

        public BorderStyle BorderTop
        {
            get { return borderFormatting.BorderTop; }
            set {
                borderFormatting.BorderTop = value;
                if (value != BorderStyle.None)
                    cfRuleRecord.IsTopBorderModified = true;
                else
                {
                    cfRuleRecord.IsTopBorderModified = (false);
                }
            }
        }

        public short BottomBorderColor
        {
            get{return borderFormatting.BottomBorderColor;}
            set
            {
                borderFormatting.BottomBorderColor=(value);
                if (value != 0)
                {
                    cfRuleRecord.IsBottomBorderModified = (true);
                }
                else
                {
                    cfRuleRecord.IsBottomBorderModified = (false);
                }
            }
        }
        public IColor BottomBorderColorColor
        {
            get
            {
                return workbook.GetCustomPalette().GetColor(borderFormatting.BottomBorderColor);
            }
            set
            {
                HSSFColor hcolor = HSSFColor.ToHSSFColor(value);
                if (hcolor == null)
                {
                    BottomBorderColor = ((short)0);
                }
                else
                {
                    BottomBorderColor = (hcolor.Indexed);
                }
            }
        }
        public short DiagonalBorderColor
        {
            get{return borderFormatting.DiagonalBorderColor;}
            set
            {
                borderFormatting.DiagonalBorderColor=(value);
                if (value != 0)
                {
                    cfRuleRecord.IsBottomLeftTopRightBorderModified = (true);
                    cfRuleRecord.IsTopLeftBottomRightBorderModified = (true);
                }
                else
                {
                    cfRuleRecord.IsBottomLeftTopRightBorderModified = (false);
                    cfRuleRecord.IsTopLeftBottomRightBorderModified = (false);
                }
            }
        }

        public IColor DiagonalBorderColorColor
        {
            get
            {
                return workbook.GetCustomPalette().GetColor(borderFormatting.DiagonalBorderColor);
            }
            set
            {
                HSSFColor hcolor = HSSFColor.ToHSSFColor(value);
                if (hcolor == null)
                {
                    DiagonalBorderColor = ((short)0);
                }
                else
                {
                    DiagonalBorderColor = (hcolor.Indexed);
                }
            }
        }

        public short LeftBorderColor
        {
            get{return borderFormatting.LeftBorderColor;}
            set
            {
                borderFormatting.LeftBorderColor=(value);
                if (value != 0)
                {
                    cfRuleRecord.IsLeftBorderModified = (true);
                }
                else
                {
                    cfRuleRecord.IsLeftBorderModified = (false);
                }
            }
        }
        public IColor LeftBorderColorColor
        {
            get
            {
                return workbook.GetCustomPalette().GetColor(borderFormatting.LeftBorderColor);
            }
            set
            {
                HSSFColor hcolor = HSSFColor.ToHSSFColor(value);
                if (hcolor == null)
                {
                    LeftBorderColor = ((short)0);
                }
                else
                {
                    LeftBorderColor = (hcolor.Indexed);
                }
            }
            
        }
        public short RightBorderColor
        {
            get{return borderFormatting.RightBorderColor;}
            set
            {
                borderFormatting.RightBorderColor=(value);
                if (value != 0)
                {
                    cfRuleRecord.IsRightBorderModified = (true);
                }
                else
                {
                    cfRuleRecord.IsRightBorderModified = (false);
                }
            }
        }
        public IColor RightBorderColorColor
        {
            get
            {
                return workbook.GetCustomPalette().GetColor(borderFormatting.RightBorderColor);
            }
            set
            {
                HSSFColor hcolor = HSSFColor.ToHSSFColor(value);
                if (hcolor == null)
                {
                    RightBorderColor = ((short)0);
                }
                else
                {
                    RightBorderColor = (hcolor.Indexed);
                }
            }
        }
        public short TopBorderColor
        {
            get{return borderFormatting.TopBorderColor;}
            set
            {
                borderFormatting.TopBorderColor=(value);
                if (value != 0)
                {
                    cfRuleRecord.IsTopBorderModified = (true);
                }
                else
                {
                    cfRuleRecord.IsTopBorderModified = (false);
                }
            }
        }
        public IColor TopBorderColorColor
        {
            get
            {
                return workbook.GetCustomPalette().GetColor(borderFormatting.TopBorderColor);
            }
            set
            {
                HSSFColor hcolor = HSSFColor.ToHSSFColor(value);
                if (hcolor == null)
                {
                    TopBorderColor = ((short)0);
                }
                else
                {
                    TopBorderColor = (hcolor.Indexed);
                }
            }
        }
        public bool IsBackwardDiagonalOn
        {
            get{return borderFormatting.IsBackwardDiagonalOn;}
            set
            {
                borderFormatting.IsBackwardDiagonalOn=value;
                if (value)
                {
                    cfRuleRecord.IsTopLeftBottomRightBorderModified = (value);
                }
            }
        }

        public bool IsForwardDiagonalOn
        {
            get { return borderFormatting.IsForwardDiagonalOn; }
            set
            {
                borderFormatting.IsForwardDiagonalOn=(value);
                if (value)
                {
                    cfRuleRecord.IsBottomLeftTopRightBorderModified = (value);
                }
            }
        }

    }
}