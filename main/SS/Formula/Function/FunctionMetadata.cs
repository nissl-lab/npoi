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

namespace NPOI.SS.Formula.Function
{
    using System;
    using System.Text;
    /**
     * Holds information about Excel built-in functions.
     * 
     * @author Josh Micich
     */
    public class FunctionMetadata
    {

        private int _index;
        private String _name;
        private int _minParams;
        private int _maxParams;
        private byte _returnClassCode;
        private byte[] _parameterClassCodes;
        private const short FUNCTION_MAX_PARAMS = 30;
        /* package */
        internal FunctionMetadata(int index, String name, int minParams, int maxParams,
            byte returnClassCode, byte[] parameterClassCodes)
        {
            _index = index;
            _name = name;
            _minParams = minParams;
            _maxParams = maxParams;
            _returnClassCode = returnClassCode;
            _parameterClassCodes = parameterClassCodes;
        }
        public int Index
        {
            get { return _index; }
        }
        public String Name
        {
            get { return _name; }
        }
        public int MinParams
        {
            get { return _minParams; }
        }
        public int MaxParams
        {
            get{return _maxParams;}
        }
        public bool HasFixedArgsLength
        {
            get { return _minParams == _maxParams; }
        }
        public byte ReturnClassCode
        {
            get { return _returnClassCode; }
        }
        public byte[] ParameterClassCodes
        {
            get { return (byte[])_parameterClassCodes.Clone(); }
        }
        public bool HasUnlimitedVarags
        {
            get
            {
                return FUNCTION_MAX_PARAMS == _maxParams;
            }
        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append(_index).Append(" ").Append(_name);
            sb.Append("]");
            return sb.ToString();
        }
    }
}