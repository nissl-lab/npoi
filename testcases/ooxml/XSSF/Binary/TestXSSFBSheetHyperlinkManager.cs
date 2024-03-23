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

namespace TestCases.XSSF.Binary
{
    using NPOI.OpenXml4Net.OPC;
    using NPOI.SS.Util;
    using NPOI.XSSF.Binary;
    using NPOI.XSSF.EventUserModel;
    using NUnit.Framework;
    using System.Collections.Generic;
    [TestFixture]
    public class TestXSSFBSheetHyperlinkManager
    {

        private static POIDataSamples _ssTests = POIDataSamples.GetSpreadSheetInstance();

        [Test]
        public void TestBasic()
        {
            OPCPackage pkg = OPCPackage.Open(_ssTests.OpenResourceAsStream("hyperlink.xlsb"));
            XSSFBReader reader = new XSSFBReader(pkg);
            XSSFReader.SheetIterator it = (XSSFReader.SheetIterator) reader.GetSheetsData();
            it.MoveNext();
            var _ = it.Current;
            XSSFBHyperlinksTable manager = new XSSFBHyperlinksTable(it.SheetPart);
            List<XSSFHyperlinkRecord> records = manager.GetHyperLinks()[new CellAddress(0, 0)];
            Assert.IsNotNull(records);
            Assert.AreEqual(1, records.Count);
            XSSFHyperlinkRecord record = records[0];
            Assert.AreEqual("http://tika.apache.org/", record.Location);
            Assert.AreEqual("rId2", record.RelId);
        }
    }
}

