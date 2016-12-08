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

    /// <summary>
    /// High level representation for Conditional Formatting Settings
    /// @author Dmitriy Kumshayev
    /// </summary>
    public class HSSFPatternFormatting : IPatternFormatting
    {
        private CFRuleRecord cfRuleRecord;
        private PatternFormatting patternFormatting;

        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFPatternFormatting"/> class.
        /// </summary>
        /// <param name="cfRuleRecord">The cf rule record.</param>
        public HSSFPatternFormatting(CFRuleRecord cfRuleRecord)
        {
            this.cfRuleRecord = cfRuleRecord;
            this.patternFormatting = cfRuleRecord.PatternFormatting;
        }

        /// <summary>
        /// Gets the pattern formatting block.
        /// </summary>
        /// <value>The pattern formatting block.</value>
        public PatternFormatting PatternFormattingBlock
        {
            get
            {
                return patternFormatting;
            }
        }

        /// <summary>
        /// Gets or sets the color of the fill background.
        /// </summary>
        /// <value>The color of the fill background.</value>
        public short FillBackgroundColor
        {
            get
            {
                return patternFormatting.FillBackgroundColor;
            }
            set
            {
                patternFormatting.FillBackgroundColor=(value);
                if (value != 0)
                {
                    cfRuleRecord.IsPatternBackgroundColorModified=(true);
                }
            }
        }

        /// <summary>
        /// Gets or sets the color of the fill foreground.
        /// </summary>
        /// <value>The color of the fill foreground.</value>
        public short FillForegroundColor
        {
            get
            {
                return patternFormatting.FillForegroundColor;
            }
            set
            {
                patternFormatting.FillForegroundColor=(value);
                if (value != 0)
                {
                    cfRuleRecord.IsPatternColorModified=(true);
                }
            }
        }

        /// <summary>
        /// Gets or sets the fill pattern.
        /// </summary>
        /// <value>The fill pattern.</value>
        public FillPattern FillPattern
        {
            get {
                return patternFormatting.FillPattern;
            }
            set {
                patternFormatting.FillPattern = value;
                if (value != 0)
                    cfRuleRecord.IsPatternStyleModified = true;
            }
        }
    }
}
