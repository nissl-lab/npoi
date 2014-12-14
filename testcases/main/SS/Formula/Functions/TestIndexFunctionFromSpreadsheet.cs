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

namespace TestCases.SS.Formula.Functions
{

    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using System;
    using NUnit.Framework;
    using TestCases.HSSF;
    using NPOI.HSSF.UserModel;
    using System.Text;
    using NPOI.SS.Util;
    using System.IO;

    /**
     * Tests INDEX() as loaded from a Test data spreadsheet.<p/>
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestIndexFunctionFromSpreadsheet:BaseTestFunctionsFromSpreadsheet
    {
        protected override string Filename
        {
            get { return "IndexFunctionTestCaseData.xls"; }
        }
    }

}
