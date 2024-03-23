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


using NPOI.XSSF.Binary;
using NUnit.Framework;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace TestCases.XSSF.Binary
{
    using NPOI.OpenXml4Net.OPC;
    [TestFixture]
    public class TestXSSFBSharedStringsTable
    {
        private static POIDataSamples _ssTests = POIDataSamples.GetSpreadSheetInstance();

        [Test]
        public void TestBasic()
        {
            OPCPackage pkg = OPCPackage.Open(_ssTests.OpenResourceAsStream("51519.xlsb"));
            List<PackagePart> parts = pkg.GetPartsByName(new Regex("/xl/sharedStrings.bin", RegexOptions.Compiled));
            Assert.AreEqual(1, parts.Count);

            XSSFBSharedStringsTable rtbl = new XSSFBSharedStringsTable(parts[0]);
            List<String> strings = rtbl.GetItems();
            Assert.AreEqual(49, strings.Count);

            Assert.AreEqual("\u30B3\u30E1\u30F3\u30C8", rtbl.GetEntryAt(0));
            Assert.AreEqual("\u65E5\u672C\u30AA\u30E9\u30AF\u30EB", rtbl.GetEntryAt(3));
            Assert.AreEqual(55, rtbl.GetCount());
            Assert.AreEqual(49, rtbl.GetUniqueCount());

            //TODO: add in tests for phonetic runs

        }
    }
}

