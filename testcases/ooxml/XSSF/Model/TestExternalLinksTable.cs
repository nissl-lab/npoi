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

namespace TestCases.XSSF.Model
{

    using NPOI.SS.UserModel;
    using NPOI.XSSF;
    using NPOI.XSSF.Model;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;


    [TestFixture]
    public class TestExternalLinksTable
    {
        [Test]
        public void None()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("SampleSS.xlsx");
            ClassicAssert.IsNotNull(wb.ExternalLinksTable);
            ClassicAssert.AreEqual(0, wb.ExternalLinksTable.Count);
        }

        [Test]
        public void BasicRead()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("ref-56737.xlsx");
            ClassicAssert.IsNotNull(wb.ExternalLinksTable);
            IName name = null;

            ClassicAssert.AreEqual(1, wb.ExternalLinksTable.Count);

            ExternalLinksTable links = wb.ExternalLinksTable[0];

            ClassicAssert.AreEqual(3, links.SheetNames.Count);
            ClassicAssert.AreEqual(2, links.DefinedNames.Count);

            ClassicAssert.AreEqual("Uses", links.SheetNames[(0)]);
            ClassicAssert.AreEqual("Defines", links.SheetNames[(1)]);
            ClassicAssert.AreEqual("56737", links.SheetNames[(2)]);

            name = links.DefinedNames[(0)];
            ClassicAssert.AreEqual("NR_Global_B2", name.NameName);
            ClassicAssert.AreEqual(-1, name.SheetIndex);
            ClassicAssert.AreEqual(null, name.SheetName);
            ClassicAssert.AreEqual("'Defines'!$B$2", name.RefersToFormula);

            name = links.DefinedNames[(1)];
            ClassicAssert.AreEqual("NR_To_A1", name.NameName);
            ClassicAssert.AreEqual(1, name.SheetIndex);
            ClassicAssert.AreEqual("Defines", name.SheetName);
            ClassicAssert.AreEqual("'Defines'!$A$1", name.RefersToFormula);

            ClassicAssert.AreEqual("56737.xlsx", links.LinkedFileName);
        }

        [Test]
        public void BasicReadWriteRead()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("ref-56737.xlsx");
            IName name = wb.ExternalLinksTable[0].DefinedNames[(1)];
            name.NameName = (/*setter*/"Testing");
            name.RefersToFormula = (/*setter*/"$A$1");

            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            ClassicAssert.AreEqual(1, wb.ExternalLinksTable.Count);
            ExternalLinksTable links = wb.ExternalLinksTable[(0)];

            name = links.DefinedNames[(0)];
            ClassicAssert.AreEqual("NR_Global_B2", name.NameName);
            ClassicAssert.AreEqual(-1, name.SheetIndex);
            ClassicAssert.AreEqual(null, name.SheetName);
            ClassicAssert.AreEqual("'Defines'!$B$2", name.RefersToFormula);

            name = links.DefinedNames[(1)];
            ClassicAssert.AreEqual("Testing", name.NameName);
            ClassicAssert.AreEqual(1, name.SheetIndex);
            ClassicAssert.AreEqual("Defines", name.SheetName);
            ClassicAssert.AreEqual("$A$1", name.RefersToFormula);
        }

        [Test]
        public void readWithReferencesToTwoExternalBooks() {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("ref2-56737.xlsx");

            ClassicAssert.IsNotNull(wb.ExternalLinksTable);
            IName name = null;

            ClassicAssert.AreEqual(2, wb.ExternalLinksTable.Count);

            // Check the first one, links to 56737.xlsx
            ExternalLinksTable links = wb.ExternalLinksTable[0];
            ClassicAssert.AreEqual("56737.xlsx", links.LinkedFileName);
            ClassicAssert.AreEqual(3, links.SheetNames.Count);
            ClassicAssert.AreEqual(2, links.DefinedNames.Count);

            ClassicAssert.AreEqual("Uses", links.SheetNames[0]);
            ClassicAssert.AreEqual("Defines", links.SheetNames[1]);
            ClassicAssert.AreEqual("56737", links.SheetNames[2]);

            name = links.DefinedNames[0];
            ClassicAssert.AreEqual("NR_Global_B2", name.NameName);
            ClassicAssert.AreEqual(-1, name.SheetIndex);
            ClassicAssert.AreEqual(null, name.SheetName);
            ClassicAssert.AreEqual("'Defines'!$B$2", name.RefersToFormula);

            name = links.DefinedNames[1];
            ClassicAssert.AreEqual("NR_To_A1", name.NameName);
            ClassicAssert.AreEqual(1, name.SheetIndex);
            ClassicAssert.AreEqual("Defines", name.SheetName);
            ClassicAssert.AreEqual("'Defines'!$A$1", name.RefersToFormula);


            // Check the second one, links to 56737.xls, slightly differently
            links = wb.ExternalLinksTable[1];
            ClassicAssert.AreEqual("56737.xls", links.LinkedFileName);
            ClassicAssert.AreEqual(2, links.SheetNames.Count);
            ClassicAssert.AreEqual(2, links.DefinedNames.Count);

            ClassicAssert.AreEqual("Uses", links.SheetNames[0]);
            ClassicAssert.AreEqual("Defines", links.SheetNames[1]);

            name = links.DefinedNames[0];
            ClassicAssert.AreEqual("NR_Global_B2", name.NameName);
            ClassicAssert.AreEqual(-1, name.SheetIndex);
            ClassicAssert.AreEqual(null, name.SheetName);
            ClassicAssert.AreEqual("'Defines'!$B$2", name.RefersToFormula);

            name = links.DefinedNames[1];
            ClassicAssert.AreEqual("NR_To_A1", name.NameName);
            ClassicAssert.AreEqual(1, name.SheetIndex);
            ClassicAssert.AreEqual("Defines", name.SheetName);
            ClassicAssert.AreEqual("'Defines'!$A$1", name.RefersToFormula);
        }
    }

}
