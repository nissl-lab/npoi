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

using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using System.Collections.Generic;
using NPOI.SS.Util;
using System;
namespace NPOI.XSSF.UserModel
{



    /**
     * @author Yegor Kozlov
     */
    public class XSSFConditionalFormatting : IConditionalFormatting
    {
        private CT_ConditionalFormatting _cf;
        private XSSFSheet _sh;

        /*package*/
        internal XSSFConditionalFormatting(XSSFSheet sh)
        {
            _cf = new CT_ConditionalFormatting();
            _sh = sh;
        }

        /*package*/
        internal XSSFConditionalFormatting(XSSFSheet sh, CT_ConditionalFormatting cf)
        {
            _cf = cf;
            _sh = sh;
        }

        /*package*/
        internal CT_ConditionalFormatting GetCTConditionalFormatting()
        {
            return _cf;
        }

        /**
          * @return array of <tt>CellRangeAddress</tt>s. Never <code>null</code>
          */
        public CellRangeAddress[] GetFormattingRanges()
        {
            List<CellRangeAddress> lst = new List<CellRangeAddress>();
            String[] regions = _cf.sqref.Split(new char[] { ' ' });
            for (int i = 0; i < regions.Length; i++)
            {
                lst.Add(CellRangeAddress.ValueOf(regions[i]));
            }
            return lst.ToArray();
        }

        /**
         * Replaces an existing Conditional Formatting rule at position idx.
         * Excel allows to create up to 3 Conditional Formatting rules.
         * This method can be useful to modify existing  Conditional Formatting rules.
         *
         * @param idx position of the rule. Should be between 0 and 2.
         * @param cfRule - Conditional Formatting rule
         */
        public void SetRule(int idx, IConditionalFormattingRule cfRule)
        {
            XSSFConditionalFormattingRule xRule = (XSSFConditionalFormattingRule)cfRule;
            _cf.GetCfRuleArray(idx).Set(xRule.GetCTCfRule());
        }

        /**
         * Add a Conditional Formatting rule.
         * Excel allows to create up to 3 Conditional Formatting rules.
         *
         * @param cfRule - Conditional Formatting rule
         */
        public void AddRule(IConditionalFormattingRule cfRule)
        {
            XSSFConditionalFormattingRule xRule = (XSSFConditionalFormattingRule)cfRule;
            _cf.AddNewCfRule().Set(xRule.GetCTCfRule());
        }

        /**
         * @return the Conditional Formatting rule at position idx.
         */
        public IConditionalFormattingRule GetRule(int idx)
        {
            return new XSSFConditionalFormattingRule(_sh, _cf.GetCfRuleArray(idx));
        }

        /**
         * @return number of Conditional Formatting rules.
         */
        public int NumberOfRules
        {
            get
            {
                return _cf.sizeOfCfRuleArray();
            }
        }
    }
}

