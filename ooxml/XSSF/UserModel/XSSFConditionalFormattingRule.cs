/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

using EnumsNET;
using NPOI.OOXML.XSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using NPOI.XSSF.Util;
using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using ConditionTypeClass = NPOI.SS.UserModel.ConditionType;

namespace NPOI.XSSF.UserModel
{

    /**
     * @author Yegor Kozlov
     */
    public class XSSFConditionalFormattingRule : IConditionalFormattingRule
    {
        private readonly CT_CfRule _cfRule;
        private readonly XSSFSheet _sh;
        
        private static readonly Dictionary<ST_CfType, ConditionType> typeLookup = new Dictionary<ST_CfType, ConditionType>();
        private static readonly Dictionary<ST_CfType, ConditionFilterType> filterTypeLookup = new Dictionary<ST_CfType, ConditionFilterType>();
        static XSSFConditionalFormattingRule()
        {
            typeLookup.Add(ST_CfType.cellIs, ConditionTypeClass.CellValueIs);
            typeLookup.Add(ST_CfType.expression, ConditionTypeClass.Formula);
            typeLookup.Add(ST_CfType.colorScale, ConditionTypeClass.ColorScale);
            typeLookup.Add(ST_CfType.dataBar, ConditionTypeClass.DataBar);
            typeLookup.Add(ST_CfType.iconSet, ConditionTypeClass.IconSet);

            // These are all subtypes of Filter, we think...
            typeLookup.Add(ST_CfType.top10, ConditionTypeClass.Filter);
            typeLookup.Add(ST_CfType.uniqueValues, ConditionTypeClass.Filter);
            typeLookup.Add(ST_CfType.duplicateValues, ConditionTypeClass.Filter);
            typeLookup.Add(ST_CfType.containsText, ConditionTypeClass.Filter);
            typeLookup.Add(ST_CfType.notContainsText, ConditionTypeClass.Filter);
            typeLookup.Add(ST_CfType.beginsWith, ConditionTypeClass.Filter);
            typeLookup.Add(ST_CfType.endsWith, ConditionTypeClass.Filter);
            typeLookup.Add(ST_CfType.containsBlanks, ConditionTypeClass.Filter);
            typeLookup.Add(ST_CfType.notContainsBlanks, ConditionTypeClass.Filter);
            typeLookup.Add(ST_CfType.containsErrors, ConditionTypeClass.Filter);
            typeLookup.Add(ST_CfType.notContainsErrors, ConditionTypeClass.Filter);
            typeLookup.Add(ST_CfType.timePeriod, ConditionTypeClass.Filter);
            typeLookup.Add(ST_CfType.aboveAverage, ConditionTypeClass.Filter);

            filterTypeLookup.Add(ST_CfType.top10, SS.UserModel.ConditionFilterType.TOP_10);
            filterTypeLookup.Add(ST_CfType.uniqueValues, SS.UserModel.ConditionFilterType.UNIQUE_VALUES);
            filterTypeLookup.Add(ST_CfType.duplicateValues, SS.UserModel.ConditionFilterType.DUPLICATE_VALUES);
            filterTypeLookup.Add(ST_CfType.containsText, SS.UserModel.ConditionFilterType.CONTAINS_TEXT);
            filterTypeLookup.Add(ST_CfType.notContainsText, SS.UserModel.ConditionFilterType.NOT_CONTAINS_TEXT);
            filterTypeLookup.Add(ST_CfType.beginsWith, SS.UserModel.ConditionFilterType.BEGINS_WITH);
            filterTypeLookup.Add(ST_CfType.endsWith, SS.UserModel.ConditionFilterType.ENDS_WITH);
            filterTypeLookup.Add(ST_CfType.containsBlanks, SS.UserModel.ConditionFilterType.CONTAINS_BLANKS);
            filterTypeLookup.Add(ST_CfType.notContainsBlanks, SS.UserModel.ConditionFilterType.NOT_CONTAINS_BLANKS);
            filterTypeLookup.Add(ST_CfType.containsErrors, SS.UserModel.ConditionFilterType.CONTAINS_ERRORS);
            filterTypeLookup.Add(ST_CfType.notContainsErrors, SS.UserModel.ConditionFilterType.NOT_CONTAINS_ERRORS);
            filterTypeLookup.Add(ST_CfType.timePeriod, SS.UserModel.ConditionFilterType.TIME_PERIOD);
            filterTypeLookup.Add(ST_CfType.aboveAverage, SS.UserModel.ConditionFilterType.ABOVE_AVERAGE);
        }
    /*package*/
    public XSSFConditionalFormattingRule(XSSFSheet sh)
        {
            _cfRule = new CT_CfRule();
            _sh = sh;
        }

        /*package*/
        internal XSSFConditionalFormattingRule(XSSFSheet sh, CT_CfRule cfRule)
        {
            _cfRule = cfRule;
            _sh = sh;
        }

        /*package*/
        internal CT_CfRule GetCTCfRule()
        {
            return _cfRule;
        }

        /*package*/
        internal CT_Dxf GetDxf(bool create)
        {
            StylesTable styles = ((XSSFWorkbook)_sh.Workbook).GetStylesSource();
            CT_Dxf dxf = null;
            if (styles.DXfsSize > 0 && _cfRule.IsSetDxfId())
            {
                int dxfId = (int)_cfRule.dxfId;
                dxf = styles.GetDxfAt(dxfId);
            }
            if (create && dxf == null)
            {
                dxf = new CT_Dxf();
                int dxfId = styles.PutDxf(dxf);
                _cfRule.dxfId = (uint)(dxfId - 1);
            }
            return dxf;
        }

        /**
         * Create a new border formatting structure if it does not exist,
         * otherwise just return existing object.
         *
         * @return - border formatting object, never returns <code>null</code>.
         */
        public IBorderFormatting CreateBorderFormatting()
        {
            CT_Dxf dxf = GetDxf(true);
            CT_Border border;
            if (!dxf.IsSetBorder())
            {
                border = dxf.AddNewBorder();
            }
            else
            {
                border = dxf.border;
            }

            return new XSSFBorderFormatting(border);
        }

        /**
         * @return - border formatting object  if defined,  <code>null</code> otherwise
         */
        public IBorderFormatting BorderFormatting
        {
            get
            {
                CT_Dxf dxf = GetDxf(false);
                if (dxf == null || !dxf.IsSetBorder()) return null;

                return new XSSFBorderFormatting(dxf.border);
            }
        }

        /**
         * Create a new font formatting structure if it does not exist,
         * otherwise just return existing object.
         *
         * @return - font formatting object, never returns <code>null</code>.
         */
        public IFontFormatting CreateFontFormatting()
        {
            CT_Dxf dxf = GetDxf(true);
            CT_Font font;
            if (!dxf.IsSetFont())
            {
                font = dxf.AddNewFont();
            }
            else
            {
                font = dxf.font;
            }

            return new XSSFFontFormatting(font);
        }

        /**
         * @return - font formatting object  if defined,  <code>null</code> otherwise
         */
        public IFontFormatting FontFormatting
        {
            get
            {
                CT_Dxf dxf = GetDxf(false);
                if (dxf == null || !dxf.IsSetFont()) return null;

                return new XSSFFontFormatting(dxf.font);
            }
        }

        /**
         * Create a new pattern formatting structure if it does not exist,
         * otherwise just return existing object.
         *
         * @return - pattern formatting object, never returns <code>null</code>.
         */
        public IPatternFormatting CreatePatternFormatting()
        {
            CT_Dxf dxf = GetDxf(true);
            CT_Fill fill;
            if (!dxf.IsSetFill())
            {
                fill = dxf.AddNewFill();
            }
            else
            {
                fill = dxf.fill;
            }

            return new XSSFPatternFormatting(fill);
        }

        /**
         * @return - pattern formatting object  if defined,  <code>null</code> otherwise
         */
        public IPatternFormatting PatternFormatting
        {
            get
            {
                CT_Dxf dxf = GetDxf(false);
                if (dxf == null || !dxf.IsSetFill()) return null;

                return new XSSFPatternFormatting(dxf.fill);
            }
        }

        public XSSFDataBarFormatting CreateDataBarFormatting(XSSFColor color)
        {
            // Is it already there?
            if (_cfRule.IsSetDataBar() && _cfRule.type == ST_CfType.dataBar)
                return DataBarFormatting as XSSFDataBarFormatting;

            // Mark it as being a Data Bar
            _cfRule.type = ST_CfType.dataBar;

            // Ensure the right element
            CT_DataBar bar = null;
            if (_cfRule.IsSetDataBar())
            {
                bar = _cfRule.dataBar;
            }
            else
            {
                bar = _cfRule.AddNewDataBar();
            }
            // Set the color
            bar.color = (color.GetCTColor());

            // Add the default thresholds
            CT_Cfvo min = bar.AddNewCfvo();
            min.type = (ST_CfvoType)Enum.Parse(typeof(ST_CfvoType), RangeType.MIN.name);
            CT_Cfvo max = bar.AddNewCfvo();
            max.type = (ST_CfvoType)Enum.Parse(typeof(ST_CfvoType), RangeType.MAX.name);

            // Wrap and return
            return new XSSFDataBarFormatting(bar);
        }
        public IDataBarFormatting DataBarFormatting
        {
            get
            {
                if (_cfRule.IsSetDataBar())
                {
                    CT_DataBar bar = _cfRule.dataBar;
                    return new XSSFDataBarFormatting(bar);
                }
                else
                {
                    return null;
                }
            }
            
        }

        public XSSFIconMultiStateFormatting CreateMultiStateFormatting(IconSet iconSet)
        {
            // Is it already there?
            if (_cfRule.IsSetIconSet() && _cfRule.type == ST_CfType.iconSet)
                return MultiStateFormatting as XSSFIconMultiStateFormatting;

            // Mark it as being an Icon Set
            _cfRule.type = (ST_CfType.iconSet);

            // Ensure the right element
            CT_IconSet icons = null;
            if (_cfRule.IsSetIconSet())
            {
                icons = _cfRule.iconSet;
            }
            else
            {
                icons = _cfRule.AddNewIconSet();
            }
            // Set the type of the icon set
            if (iconSet.name != null)
            {
                ST_IconSetType xIconSet =Enums.Parse<ST_IconSetType>(iconSet.name, false, EnumFormat.Description);
                icons.iconSet = xIconSet;
            }

            // Add a default set of thresholds
            int jump = 100 / iconSet.num;
            ST_CfvoType type = (ST_CfvoType)Enum.Parse(typeof(ST_CfvoType), RangeType.PERCENT.name);
            for (int i = 0; i < iconSet.num; i++)
            {
                CT_Cfvo cfvo = icons.AddNewCfvo();
                cfvo.type = (type);
                cfvo.val = (i * jump).ToString();
            }

            // Wrap and return
            return new XSSFIconMultiStateFormatting(icons);
        }
        public IIconMultiStateFormatting MultiStateFormatting
        {
            get
            {
                if (_cfRule.IsSetIconSet())
                {
                    CT_IconSet icons = _cfRule.iconSet;
                    return new XSSFIconMultiStateFormatting(icons);
                }
                else
                {
                    return null;
                }
            }
        }

        public XSSFColorScaleFormatting CreateColorScaleFormatting()
        {
            // Is it already there?
            if (_cfRule.IsSetColorScale() && _cfRule.type == ST_CfType.colorScale)
                return ColorScaleFormatting as XSSFColorScaleFormatting;

            // Mark it as being a Color Scale
            _cfRule.type = (ST_CfType.colorScale);

            // Ensure the right element
            CT_ColorScale scale = null;
            if (_cfRule.IsSetColorScale())
            {
                scale = _cfRule.colorScale;
            }
            else
            {
                scale = _cfRule.AddNewColorScale();
            }

            // Add a default set of thresholds and colors
            if (scale.SizeOfCfvoArray() == 0)
            {
                CT_Cfvo cfvo;
                cfvo = scale.AddNewCfvo();
                cfvo.type = (ST_CfvoType)Enum.Parse(typeof(ST_CfvoType), RangeType.MIN.name);
                cfvo = scale.AddNewCfvo();
                cfvo.type = (ST_CfvoType)Enum.Parse(typeof(ST_CfvoType), RangeType.PERCENTILE.name);
                cfvo.val = ("50");
                cfvo = scale.AddNewCfvo();
                cfvo.type = (ST_CfvoType)Enum.Parse(typeof(ST_CfvoType), RangeType.MAX.name);

                for (int i = 0; i < 3; i++)
                {
                    scale.AddNewColor();
                }
            }

            // Wrap and return
            return new XSSFColorScaleFormatting(scale);
        }
        public IColorScaleFormatting ColorScaleFormatting
        {
            get
            {
                if (_cfRule.IsSetColorScale())
                {
                    CT_ColorScale scale = _cfRule.colorScale;
                    return new XSSFColorScaleFormatting(scale);
                }
                else
                {
                    return null;
                }
            }
            
        }
        /**
         * Type of conditional formatting rule.
         */
        public ConditionType ConditionType
        {
            get
            {
                return typeLookup[(_cfRule.type)];
            }
            
        }
        /**
         * The comparison function used when the type of conditional formatting is Set to
         * {@link ConditionalFormattingRule#CONDITION_TYPE_CELL_VALUE_IS}
         * <p>
         *     MUST be a constant from {@link NPOI.ss.usermodel.ComparisonOperator}
         * </p>
         *
         * @return the conditional format operator
         */
        public ComparisonOperator ComparisonOperation
        {
            get
            {
                ST_ConditionalFormattingOperator? op = _cfRule.@operator;
                if (op == null) 
                    return ComparisonOperator.NoComparison;

                switch (op)
                {
                    case ST_ConditionalFormattingOperator.lessThan: return ComparisonOperator.LessThan;
                    case ST_ConditionalFormattingOperator.lessThanOrEqual: return ComparisonOperator.LessThanOrEqual;
                    case ST_ConditionalFormattingOperator.greaterThan: return ComparisonOperator.GreaterThan;
                    case ST_ConditionalFormattingOperator.greaterThanOrEqual: return ComparisonOperator.GreaterThanOrEqual;
                    case ST_ConditionalFormattingOperator.equal: return ComparisonOperator.Equal;
                    case ST_ConditionalFormattingOperator.notEqual: return ComparisonOperator.NotEqual;
                    case ST_ConditionalFormattingOperator.between: return ComparisonOperator.Between;
                    case ST_ConditionalFormattingOperator.notBetween: return ComparisonOperator.NotBetween;
                }
                return ComparisonOperator.NoComparison;
            }
        }

        /**
         * The formula used to Evaluate the first operand for the conditional formatting rule.
         * <p>
         * If the condition type is {@link ConditionalFormattingRule#CONDITION_TYPE_CELL_VALUE_IS},
         * this field is the first operand of the comparison.
         * If type is {@link ConditionalFormattingRule#CONDITION_TYPE_FORMULA}, this formula is used
         * to determine if the conditional formatting is applied.
         * </p>
         * <p>
         * If comparison type is {@link ConditionalFormattingRule#CONDITION_TYPE_FORMULA} the formula MUST be a Boolean function
         * </p>
         *
         * @return  the first formula
         */
        public String Formula1
        {
            get
            {
                return _cfRule.SizeOfFormulaArray() > 0 ? _cfRule.GetFormulaArray(0) : null;
            }
        }

        /**
         * The formula used to Evaluate the second operand of the comparison when
         * comparison type is  {@link ConditionalFormattingRule#CONDITION_TYPE_CELL_VALUE_IS} and operator
         * is either {@link NPOI.ss.usermodel.ComparisonOperator#BETWEEN} or {@link NPOI.ss.usermodel.ComparisonOperator#NOT_BETWEEN}
         *
         * @return  the second formula
         */
        public String Formula2
        {
            get
            {
                return _cfRule.SizeOfFormulaArray() == 2 ? _cfRule.GetFormulaArray(1) : null;
            }
        }

        public bool StopIfTrue
        {
            get
            {
                return _cfRule.stopIfTrue;
            }
        }
        public String Text
        {
            get
            {
                return _cfRule.text;
            }
        }
        public int Priority
        {
            get
            {
                int priority = _cfRule.priority;
                // priorities start at 1, if it is less, it is undefined, use definition order in caller
                return priority >= 1 ? priority : 0;
            }
        }
        public ExcelNumberFormat NumberFormat
        {
            get
            {
                CT_Dxf dxf = GetDxf(false);
                if (dxf == null || !dxf.IsSetNumFmt()) return null;

                CT_NumFmt numFmt = dxf.numFmt;
                return new ExcelNumberFormat((int)numFmt.numFmtId, numFmt.formatCode);
            }
        }

        public ConditionFilterType? ConditionFilterType
        {
            get
            {
                if (!filterTypeLookup.TryGetValue(_cfRule.type, out ConditionFilterType type))
                    return null;
                return type;
            }
        }
        public IConditionFilterData FilterConfiguration
        {
            get
            {
                return new XSSFConditionFilterData(_cfRule);
            }
        }
    }
}


