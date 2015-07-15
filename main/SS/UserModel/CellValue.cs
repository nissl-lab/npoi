/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.UserModel
{
    using System;
    using System.Text;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using System.Globalization;

    /**
     * Mimics the 'data view' of a cell. This allows formula Evaluator
     * to return a CellValue instead of precasting the value to String
     * or Number or bool type.
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     */
    public class CellValue
    {
        public static readonly CellValue TRUE = new CellValue(CellType.Boolean, 0.0, true, null, 0);
        public static readonly CellValue FALSE = new CellValue(CellType.Boolean, 0.0, false, null, 0);

        private CellType _cellType;
        private double _numberValue;
        private bool _boolValue;
        private String _textValue;
        private int _errorCode;

        private CellValue(CellType cellType, double numberValue, bool boolValue,
                String textValue, int errorCode)
        {
            _cellType = cellType;
            _numberValue = numberValue;
            _boolValue = boolValue;
            _textValue = textValue;
            _errorCode = errorCode;
        }


        public CellValue(double numberValue)
            : this(CellType.Numeric, numberValue, false, null, 0)
        {

        }
        public static CellValue ValueOf(bool boolValue)
        {
            return boolValue ? TRUE : FALSE;
        }
        public CellValue(String stringValue)
            : this(CellType.String, 0.0, false, stringValue, 0)
        {

        }
        public static CellValue GetError(int errorCode)
        {
            return new CellValue(CellType.Error, 0.0, false, null, errorCode);
        }


        /**
         * @return Returns the boolValue.
         */
        public bool BooleanValue
        {
            get
            {
                return _boolValue;
            }
        }
        /**
         * @return Returns the numberValue.
         */
        public double NumberValue
        {
            get
            {
                return _numberValue;
            }
        }
        /**
         * @return Returns the stringValue.
         */
        public String StringValue
        {
            get
            {
                return _textValue;
            }
        }
        /**
         * @return Returns the cellType.
         */
        public CellType CellType
        {
            get
            {
                return _cellType;
            }
        }
        /**
         * @return Returns the errorValue.
         */
        //the return value should be sbyte? the byte in java is signed(-128~127) and is unsiged(0~255) in c#.
        //if use byte, the test NPOI.SS.Formula.TestWorkbookEvaluator.TestResultOutsideRange failed.
        public sbyte ErrorValue
        {
            get
            {
                return (sbyte)_errorCode;
            }
        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append(FormatAsString());
            sb.Append("]");
            return sb.ToString();
        }

        public String FormatAsString()
        {
            switch (_cellType)
            {
                case CellType.Numeric:
                    string result = _numberValue.ToString(CultureInfo.InvariantCulture);
                    //if (result.IndexOf(".") < 0)
                    //    result = result + ".0";
                    return result;
                case CellType.String:
                    return '"' + _textValue + '"';
                case CellType.Boolean:
                    return _boolValue ? "TRUE" : "FALSE";
                case CellType.Error:
                    return ErrorEval.GetText(_errorCode);
            }
            return "<error unexpected cell type " + _cellType + ">";
        }
    }
}

