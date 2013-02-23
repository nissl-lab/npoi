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
        private CFRuleRecord cfRuleRecord;
        private BorderFormatting borderFormatting;

        public HSSFBorderFormatting(CFRuleRecord cfRuleRecord)
        {
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
            set {
                borderFormatting.BorderBottom = value;
                if (value != BorderStyle.None)
                    cfRuleRecord.IsBottomBorderModified = true;
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
            }
        }

        public BorderStyle BorderLeft
        {
            get { return borderFormatting.BorderLeft; }
            set {
                borderFormatting.BorderLeft = value;
                if (value != BorderStyle.None)
                    cfRuleRecord.IsLeftBorderModified = true;
            }
        }

        public BorderStyle BorderRight
        {
            get { return borderFormatting.BorderRight; }
            set {
                borderFormatting.BorderRight = value;
                if (value != BorderStyle.None)
                    cfRuleRecord.IsRightBorderModified = true;
            }
        }

        public BorderStyle BorderTop
        {
            get { return borderFormatting.BorderTop; }
            set {
                borderFormatting.BorderTop = value;
                if (value != BorderStyle.None)
                    cfRuleRecord.IsTopBorderModified = true;
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