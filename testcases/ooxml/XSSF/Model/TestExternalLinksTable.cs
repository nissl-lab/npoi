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

namespace NPOI.XSSF.Model
{
    using System;

    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;


    [TestFixture]
    public class TestExternalLinksTable
    {
        [Test]
        public void None()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("SampleSS.xlsx");
            Assert.AreEqual(null, wb.ExternalLinksTable);
        }

        [Test]
        public void BasicRead()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("ref-56737.xlsx");
            Assert.IsNotNull(wb.ExternalLinksTable);
            IName name = null;

            ExternalLinksTable links = wb.ExternalLinksTable;
            Assert.AreEqual(3, links.SheetNames.Count);
            Assert.AreEqual(2, links.DefinedNames.Count);

            Assert.AreEqual("Uses", links.SheetNames[(0)]);
            Assert.AreEqual("Defines", links.SheetNames[(1)]);
            Assert.AreEqual("56737", links.SheetNames[(2)]);

            name = links.DefinedNames[(0)];
            Assert.AreEqual("NR_Global_B2", name.NameName);
            Assert.AreEqual(-1, name.SheetIndex);
            Assert.AreEqual(null, name.SheetName);
            Assert.AreEqual("'Defines'!$B$2", name.RefersToFormula);

            name = links.DefinedNames[(1)];
            Assert.AreEqual("NR_To_A1", name.NameName);
            Assert.AreEqual(1, name.SheetIndex);
            Assert.AreEqual("Defines", name.SheetName);
            Assert.AreEqual("'Defines'!$A$1", name.RefersToFormula);
        }

        [Test]
        public void BasicReadWriteRead()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("ref-56737.xlsx");
            IName name = wb.ExternalLinksTable.DefinedNames[(1)];
            name.NameName = (/*setter*/"Testing");
            name.RefersToFormula = (/*setter*/"$A$1");

            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            ExternalLinksTable links = wb.ExternalLinksTable;

            name = links.DefinedNames[(0)];
            Assert.AreEqual("NR_Global_B2", name.NameName);
            Assert.AreEqual(-1, name.SheetIndex);
            Assert.AreEqual(null, name.SheetName);
            Assert.AreEqual("'Defines'!$B$2", name.RefersToFormula);

            name = links.DefinedNames[(1)];
            Assert.AreEqual("Testing", name.NameName);
            Assert.AreEqual(1, name.SheetIndex);
            Assert.AreEqual("Defines", name.SheetName);
            Assert.AreEqual("$A$1", name.RefersToFormula);
        }
    }

}
