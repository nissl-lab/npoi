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

using System.Collections.Generic;
using NPOI.SS.Util;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula;
using System;
using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using System.Linq;
using NPOI.SS.UserModel.Helpers;
using NPOI.OOXML.XSSF.UserModel.Helpers;

namespace NPOI.XSSF.UserModel.Helpers
{
    /**
     * @author Yegor Kozlov
     */
    public class XSSFRowShifter : RowShifter
    {

        public XSSFRowShifter(XSSFSheet sh)
            : base(sh)
        {
            sheet = sh;
        }

        public override void UpdateConditionalFormatting(FormulaShifter formulaShifter)
        {
            XSSFRowColShifter.UpdateConditionalFormatting(sheet, formulaShifter);
        }

        public override void UpdateFormulas(FormulaShifter formulaShifter)
        {
            XSSFRowColShifter.UpdateFormulas(sheet, formulaShifter);
        }

        public override void UpdateHyperlinks(FormulaShifter formulaShifter)
        {
            XSSFRowColShifter.UpdateHyperlinks(sheet, formulaShifter);
        }

        public override void UpdateNamedRanges(FormulaShifter formulaShifter)
        {
            XSSFRowColShifter.UpdateNamedRanges(sheet, formulaShifter);
        }

        public override void UpdateRowFormulas(IRow row, FormulaShifter formulaShifter)
        {
            XSSFRowColShifter.UpdateRowFormulas(row, formulaShifter);
        }
    }
}


