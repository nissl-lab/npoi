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

namespace NPOI.xssf.usermodel;

using NPOI.ss.usermodel.ConditionalFormatting;
using NPOI.ss.usermodel.ConditionalFormattingRule;
using NPOI.ss.util.CellRangeAddress;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;



/**
 * @author Yegor Kozlov
 */
public class XSSFConditionalFormatting : ConditionalFormatting {
    private CTConditionalFormatting _cf;
    private XSSFSheet _sh;

    /*package*/ XSSFConditionalFormatting(XSSFSheet sh){
        _cf = CTConditionalFormatting.Factory.newInstance();
        _sh = sh;
    }

    /*package*/ XSSFConditionalFormatting(XSSFSheet sh, CTConditionalFormatting cf){
        _cf = cf;
        _sh = sh;
    }

    /*package*/  CTConditionalFormatting GetCTConditionalFormatting(){
        return _cf;
    }

    /**
      * @return array of <tt>CellRangeAddress</tt>s. Never <code>null</code>
      */
     public CellRangeAddress[] GetFormattingRanges(){
         List<CellRangeAddress> lst = new List<CellRangeAddress>();
         foreach (Object stRef in _cf.GetSqref()) {
             String[] regions = stRef.ToString().split(" ");
             for (int i = 0; i < regions.Length; i++) {
                 lst.Add(CellRangeAddress.ValueOf(regions[i]));
             }
         }
         return lst.ToArray(new CellRangeAddress[lst.Count]);
     }

     /**
      * Replaces an existing Conditional Formatting rule at position idx.
      * Excel allows to create up to 3 Conditional Formatting rules.
      * This method can be useful to modify existing  Conditional Formatting rules.
      *
      * @param idx position of the rule. Should be between 0 and 2.
      * @param cfRule - Conditional Formatting rule
      */
     public void SetRule(int idx, ConditionalFormattingRule cfRule){
         XSSFConditionalFormattingRule xRule = (XSSFConditionalFormattingRule)cfRule;
         _cf.GetCfRuleArray(idx).Set(xRule.GetCTCfRule());
     }

     /**
      * Add a Conditional Formatting rule.
      * Excel allows to create up to 3 Conditional Formatting rules.
      *
      * @param cfRule - Conditional Formatting rule
      */
     public void AddRule(ConditionalFormattingRule cfRule){
        XSSFConditionalFormattingRule xRule = (XSSFConditionalFormattingRule)cfRule;
         _cf.AddNewCfRule().Set(xRule.GetCTCfRule());
     }

     /**
      * @return the Conditional Formatting rule at position idx.
      */
     public XSSFConditionalFormattingRule GetRule(int idx){
         return new XSSFConditionalFormattingRule(_sh, _cf.GetCfRuleArray(idx));
     }

     /**
      * @return number of Conditional Formatting rules.
      */
     public int GetNumberOfRules(){
         return _cf.sizeOfCfRuleArray();
     }
}


