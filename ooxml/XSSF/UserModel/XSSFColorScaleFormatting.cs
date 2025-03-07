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
    using System.Collections.Generic;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.SS.UserModel;

    /**
     * High level representation for Color Scale / Color Gradient Formatting 
     *  component of Conditional Formatting Settings
     */
    public class XSSFColorScaleFormatting : IColorScaleFormatting {
        readonly CT_ColorScale _scale;

        /*package*/
        public XSSFColorScaleFormatting(CT_ColorScale scale) {
            _scale = scale;
        }

        public int NumControlPoints
        {
            get { return _scale.SizeOfCfvoArray(); }
            set {
                while (value < _scale.SizeOfCfvoArray())
                {
                    _scale.RemoveCfvo(_scale.SizeOfCfvoArray() - 1);
                    _scale.RemoveColor(_scale.SizeOfColorArray() - 1);
                }
                while (value > _scale.SizeOfCfvoArray())
                {
                    _scale.AddNewCfvo();
                    _scale.AddNewColor();
                }
            }
        }

        public IColor[] Colors
        {
            get
            {
                CT_Color[] ctcols = _scale.color.ToArray();//.ColorArray;
                XSSFColor[] c = new XSSFColor[ctcols.Length];
                for (int i = 0; i < ctcols.Length; i++)
                {
                    c[i] = new XSSFColor(ctcols[i]);
                }
                return c;
            }
            set
            {
                CT_Color[] ctcols = new CT_Color[value.Length];
                for (int i = 0; i < value.Length; i++)
                {
                    ctcols[i] = ((XSSFColor)value[i]).GetCTColor();
                }
                _scale.color = new List<CT_Color>(ctcols);
            }
        }

        public IConditionalFormattingThreshold[] Thresholds
        {
            get
            {
                CT_Cfvo[] cfvos = _scale.cfvo.ToArray();
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
                _scale.cfvo = new List<CT_Cfvo>(cfvos);
            }
        }

        public XSSFColor CreateColor() {
            return new XSSFColor(_scale.AddNewColor());
        }
        public IConditionalFormattingThreshold CreateThreshold() {
            return new XSSFConditionalFormattingThreshold(_scale.AddNewCfvo());
        }
    }

}