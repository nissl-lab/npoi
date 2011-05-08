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
    using System;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.CF;

    /**
     * High level representation for Border Formatting component
     * of Conditional Formatting Settings
     * 
     * @author Dmitriy Kumshayev
     *
     */
    public class HSSFBorderFormatting
    {
        /** No border */
        public static short BORDER_NONE = BorderFormatting.BORDER_NONE;
        /** Thin border */
        public static short BORDER_THIN = BorderFormatting.BORDER_THIN;
        /** Medium border */
        public static short BORDER_MEDIUM = BorderFormatting.BORDER_MEDIUM;
        /** dash border */
        public static short BORDER_DASHED = BorderFormatting.BORDER_DASHED;
        /** dot border */
        public static short BORDER_HAIR = BorderFormatting.BORDER_HAIR;
        /** Thick border */
        public static short BORDER_THICK = BorderFormatting.BORDER_THICK;
        /** double-line border */
        public static short BORDER_DOUBLE = BorderFormatting.BORDER_DOUBLE;
        /** hair-line border */
        public static short BORDER_DOTTED = BorderFormatting.BORDER_DOTTED;
        /** Medium dashed border */
        public static short BORDER_MEDIUM_DASHED = BorderFormatting.BORDER_MEDIUM_DASHED;
        /** dash-dot border */
        public static short BORDER_DASH_DOT = BorderFormatting.BORDER_DASH_DOT;
        /** medium dash-dot border */
        public static short BORDER_MEDIUM_DASH_DOT = BorderFormatting.BORDER_MEDIUM_DASH_DOT;
        /** dash-dot-dot border */
        public static short BORDER_DASH_DOT_DOT = BorderFormatting.BORDER_DASH_DOT_DOT;
        /** medium dash-dot-dot border */
        public static short BORDER_MEDIUM_DASH_DOT_DOT = BorderFormatting.BORDER_MEDIUM_DASH_DOT_DOT;
        /** slanted dash-dot border */
        public static short BORDER_SLANTED_DASH_DOT = BorderFormatting.BORDER_SLANTED_DASH_DOT;


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

        public short BorderBottom
        {
            get{return borderFormatting.BorderBottom;}
            set
            {
                borderFormatting.BorderBottom=(value);
                if (value != 0)
                {
                    cfRuleRecord.IsBottomBorderModified = (true);
                }
            }
        }

        public short BorderDiagonal
        {
            get{return borderFormatting.BorderDiagonal;}
            set
            {
                borderFormatting.BorderDiagonal=(value);
                if (value != 0)
                {
                    cfRuleRecord.IsBottomLeftTopRightBorderModified = (true);
                    cfRuleRecord.IsTopLeftBottomRightBorderModified = (true);
                }
            }
        }

        public short BorderLeft
        {
            get{return borderFormatting.BorderLeft;}
            set
            {
                borderFormatting.BorderLeft=(value);
                if (value != 0)
                {
                    cfRuleRecord.IsLeftBorderModified = (true);
                }
            }
        }

        public short BorderRight
        {
            get{return borderFormatting.BorderRight;}
            set
            {
                borderFormatting.BorderRight=(value);
                if (value != 0)
                {
                    cfRuleRecord.IsRightBorderModified = (true);
                }
            }
        }

        public short BorderTop
        {
            get{return borderFormatting.BorderTop;}
            set
            {
                borderFormatting.BorderTop=(value);
                if (value != 0)
                {
                    cfRuleRecord.IsTopBorderModified = (true);
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