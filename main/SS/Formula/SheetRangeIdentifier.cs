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

    public class SheetRangeIdentifier : SheetIdentifier
    {
        public NameIdentifier _lastSheetIdentifier;

        public SheetRangeIdentifier(String bookName, NameIdentifier firstSheetIdentifier, NameIdentifier lastSheetIdentifier)
            : base(bookName, firstSheetIdentifier)
        {
            _lastSheetIdentifier = lastSheetIdentifier;
        }
        public NameIdentifier FirstSheetIdentifier
        {
            get
            {
                return base.SheetId;
            }
        }
        public NameIdentifier LastSheetIdentifier
        {
            get
            {
                return _lastSheetIdentifier;
            }
        }
        protected override void AsFormulaString(StringBuilder sb)
        {
            base.AsFormulaString(sb);
            sb.Append(':');
            if (_lastSheetIdentifier.IsQuoted)
            {
                sb.Append("'").Append(_lastSheetIdentifier.Name).Append("'");
            }
            else
            {
                sb.Append(_lastSheetIdentifier.Name);
            }
        }
    }
}