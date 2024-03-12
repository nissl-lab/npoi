/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using NPOI.XSSF.UserModel;
using NUnit.Framework;

namespace TestCases.XSSF.UserModel
{

    [TestFixture]
    public class TestXSSFEvaluationSheet
    {

        [Test]
        public void Test()
        {

            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet("test") as XSSFSheet;
            XSSFRow row = sheet.CreateRow(0) as XSSFRow;
            row.CreateCell(0);
            XSSFEvaluationSheet evalsheet = new XSSFEvaluationSheet(sheet);

            Assert.IsNotNull(evalsheet.GetCell(0, 0), "Cell 0,0 is found");
            Assert.IsNull(evalsheet.GetCell(0, 1), "Cell 0,1 is not found");
            Assert.IsNull(evalsheet.GetCell(1, 0), "Cell 1,0 is not found");

            // now add Cell 0,1
            row.CreateCell(1);

            Assert.IsNotNull(evalsheet.GetCell(0, 0), "Cell 0,0 is found");
            Assert.IsNotNull(evalsheet.GetCell(0, 1), "Cell 0,1 is now also found");
            Assert.IsNull(evalsheet.GetCell(1, 0), "Cell 1,0 is not found");

            // After clearing all values it also works
            row.CreateCell(2);
            evalsheet.ClearAllCachedResultValues();

            Assert.IsNotNull(evalsheet.GetCell(0, 0), "Cell 0,0 is found");
            Assert.IsNotNull(evalsheet.GetCell(0, 2), "Cell 0,2 is now also found");
            Assert.IsNull(evalsheet.GetCell(1, 0), "Cell 1,0 is not found");

            // other things
            Assert.AreEqual(sheet, evalsheet.XSSFSheet);
        }
    }
}

