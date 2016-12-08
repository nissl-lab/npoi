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
namespace NPOI.XSSF.UserModel
{
    /**
     * @author Yegor Kozlov
     */
    public class XSSFPatternFormatting : IPatternFormatting
    {
        CT_Fill _fill;

        public XSSFPatternFormatting(CT_Fill fill)
        {
            _fill = fill;
        }

        public short FillBackgroundColor
        {
            get
            {
                if (!_fill.IsSetPatternFill()) return 0;

                return _fill.GetPatternFill().bgColor.indexedSpecified ? (short)_fill.GetPatternFill().bgColor.indexed : (short)0;
            }
            set 
            {
                CT_PatternFill ptrn =
                _fill.IsSetPatternFill() ? _fill.GetPatternFill() : _fill.AddNewPatternFill();
                CT_Color bgColor = new CT_Color();
                bgColor.indexed = (uint)value;
                bgColor.indexedSpecified = true;
                ptrn.bgColor = (bgColor);
            }
        }

        public short FillForegroundColor
        {
            get
            {
                if (!_fill.IsSetPatternFill() || !_fill.GetPatternFill().IsSetFgColor())
                    return 0;

                return _fill.GetPatternFill().fgColor.indexedSpecified?  (short)_fill.GetPatternFill().fgColor.indexed : (short)0;
            }
            set 
            {
                CT_PatternFill ptrn = _fill.IsSetPatternFill() ? _fill.GetPatternFill() : _fill.AddNewPatternFill();
                CT_Color fgColor = new CT_Color();
                fgColor.indexed = (uint)(value);
                fgColor.indexedSpecified = true;
                ptrn.fgColor = (fgColor);
            }
        }

        public FillPattern FillPattern
        {
            get
            {
                if (!_fill.IsSetPatternFill() || !_fill.GetPatternFill().IsSetPatternType())
                    return 0;

                return (FillPattern) _fill.GetPatternFill().patternType;
            }
            set 
            {
                CT_PatternFill ptrn = _fill.IsSetPatternFill() ? _fill.GetPatternFill() : _fill.AddNewPatternFill();
                ptrn.patternType = (ST_PatternType)(value);
            }
        }

    }

}

