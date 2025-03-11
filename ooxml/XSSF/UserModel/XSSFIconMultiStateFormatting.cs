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
namespace NPOI.XSSF.UserModel
{
    using System;
    using EnumsNET;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.Util;

    /**
     * High level representation for Icon / Multi-State Formatting 
     *  component of Conditional Formatting Settings
     */
    public class XSSFIconMultiStateFormatting : IIconMultiStateFormatting
    {
        readonly CT_IconSet _iconset;

        /*package*/
        internal XSSFIconMultiStateFormatting(CT_IconSet iconset)
        {
            _iconset = iconset;
        }

        public IconSet IconSet
        {
            get
            {
                String set = _iconset.iconSet.ToString();
                return IconSet.ByOOXMLName(set);
            }
            set
            {
                ST_IconSetType xIconSet = Enums.Parse<ST_IconSetType>(value.name, false, EnumFormat.Description);
                _iconset.iconSet = (xIconSet);
            }
        }

        public bool IsIconOnly
        {
            get
            {
                if (_iconset.IsSetShowValue())
                    return !_iconset.showValue;
                return false;
            }
            set
            {
                _iconset.showValue = !value;
            }
        }

        public bool IsReversed
        {
            get
            {
                if (_iconset.reverse)
                    return _iconset.reverse;
                return false;
            }
            set
            {
                _iconset.reverse = (value);
            }
        }

        public IConditionalFormattingThreshold[] Thresholds
        {
            get
            {
                CT_Cfvo[] cfvos = _iconset.cfvo.ToArray();
                XSSFConditionalFormattingThreshold[] t =
                        new XSSFConditionalFormattingThreshold[cfvos.Length];
                for (int i = 0; i < cfvos.Length; i++)
                {
                    t[i] = new XSSFConditionalFormattingThreshold(cfvos[i]);
                }
                return t;
            }
            set
            {
                CT_Cfvo[] cfvos = new CT_Cfvo[value.Length];
                for (int i = 0; i < value.Length; i++)
                {
                    cfvos[i] = ((XSSFConditionalFormattingThreshold)value[i]).CTCfvo;
                }
                _iconset.cfvo = new System.Collections.Generic.List<CT_Cfvo>(cfvos);
            }
        }

        public IConditionalFormattingThreshold CreateThreshold()
        {
            return new XSSFConditionalFormattingThreshold(_iconset.AddNewCfvo());
        }
    }

}