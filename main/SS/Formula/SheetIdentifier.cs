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

namespace NPOI.SS.Formula
{
    using System;
    using System.Text;

    public class SheetIdentifier
    {
        public String _bookName;
        public NameIdentifier _sheetIdentifier;

        public SheetIdentifier(String bookName, NameIdentifier sheetIdentifier)
        {
            _bookName = bookName;
            _sheetIdentifier = sheetIdentifier;
        }
        public String BookName
        {
            get
            {
                return _bookName;
            }
        }
        public NameIdentifier SheetId
        {
            get
            {
                return _sheetIdentifier;
            }
        }
        protected virtual void AsFormulaString(StringBuilder sb)
        {
            if (_bookName != null)
            {
                sb.Append(" [").Append(_sheetIdentifier.Name).Append("]");
            }
            if (_sheetIdentifier.IsQuoted)
            {
                sb.Append("'").Append(_sheetIdentifier.Name).Append("'");
            }
            else
            {
                sb.Append(_sheetIdentifier.Name);
            }
        }
        public String AsFormulaString()
        {
            StringBuilder sb = new StringBuilder(32);
            AsFormulaString(sb);
            return sb.ToString();
        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(this.GetType().Name);
            sb.Append(" [");
            AsFormulaString(sb);
            sb.Append("]");
            return sb.ToString();
        }
    }
}