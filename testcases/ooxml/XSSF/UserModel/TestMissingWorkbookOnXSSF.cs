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

namespace NPOI.XSSF.UserModel
{
    using System;

    using NPOI.HSSF;
    using NPOI.SS.Formula;
    using NPOI.XSSF;
    using TestCases.SS.Formula;
    using NUnit.Framework;
    using TestCases.HSSF;

    /**
     * XSSF Specific version of the Missing Workbooks test
     */
    [TestFixture]
    public class TestMissingWorkbookOnXSSF : TestMissingWorkbook
    {
        public TestMissingWorkbookOnXSSF()
            : base("52575_main.xlsx", "source_dummy.xlsx", "52575_source.xls")
        {
            ;
        }

        [SetUp]
        protected override void SetUp()
        {
            mainWorkbook = XSSFTestDataSamples.OpenSampleWorkbook(this.MAIN_WORKBOOK_FILENAME);
            sourceWorkbook = HSSFTestDataSamples.OpenSampleWorkbook(this.SOURCE_WORKBOOK_FILENAME);

            Assert.IsNotNull(mainWorkbook);
            Assert.IsNotNull(sourceWorkbook);
        }
    }
}