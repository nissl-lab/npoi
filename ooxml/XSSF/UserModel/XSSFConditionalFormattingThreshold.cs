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
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.SS.UserModel;
    using System;


    /**
     * High level representation for Icon / Multi-State / Databar /
     *  Colour Scale change thresholds
     */
    public class XSSFConditionalFormattingThreshold : IConditionalFormattingThreshold
    {
        private CT_Cfvo cfvo;

        protected internal XSSFConditionalFormattingThreshold(CT_Cfvo cfvo)
        {
            this.cfvo = cfvo;
        }

        protected internal CT_Cfvo CTCfvo
        {
            get { return cfvo; }
        }

        public RangeType RangeType
        {
            get
            {
                return RangeType.ByName(cfvo.type.ToString());
            }
            set
            {
                ST_CfvoType xtype = (ST_CfvoType)Enum.Parse(typeof(ST_CfvoType), value.name);
                cfvo.type = (/*setter*/xtype);
            }
        }

        public String Formula
        {
            get
            {
                if (cfvo.type == ST_CfvoType.formula)
                {
                    return cfvo.val;
                }
                return null;
            }
            set
            {
                cfvo.val = value;
            }
        }

        public double? Value
        {
            get
            {
                if (cfvo.type == ST_CfvoType.formula ||
                cfvo.type == ST_CfvoType.min ||
                cfvo.type == ST_CfvoType.max)
                {
                    return null;
                }
                if (cfvo.IsSetVal())
                {
                    return Double.Parse(cfvo.val);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if (value == null)
                {
                    cfvo.UnsetVal();
                }
                else
                {
                    cfvo.val = value.ToString();
                }
            }
        }
    }

}