/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.XSSF.EventUserModel
{

    using System.Text.RegularExpressions;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.XSSF.Model;
    using NPOI.XSSF.UserModel;
    using NPOI.XSSF.EventUserModel;
    using NUnit.Framework;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NUnit.Framework.Legacy;

    /// <summary>
    /// Tests for <see cref="XSSFReader" />
    /// </summary>
    [TestFixture]
    public sealed class TestReadOnlySharedStringsTable
    {
        private static POIDataSamples _ssTests = POIDataSamples.GetSpreadSheetInstance();

        [Test]
        public void TestParse()
        {

            OPCPackage pkg = OPCPackage.Open(_ssTests.OpenResourceAsStream("SampleSS.xlsx"));
            List<PackagePart> parts = pkg.GetPartsByName(new Regex("/xl/sharedStrings.xml", RegexOptions.Compiled));
            ClassicAssert.AreEqual(1, parts.Count);

            SharedStringsTable stbl = new SharedStringsTable(parts[0]);
            ReadOnlySharedStringsTable rtbl = new ReadOnlySharedStringsTable(parts[0]);

            ClassicAssert.AreEqual(stbl.Count, rtbl.Count);
            ClassicAssert.AreEqual(stbl.UniqueCount, rtbl.UniqueCount);

            ClassicAssert.AreEqual(stbl.Items.Count, stbl.UniqueCount);
            ClassicAssert.AreEqual(rtbl.Items.Count, rtbl.UniqueCount);
            for(int i = 0; i < stbl.UniqueCount; i++)
            {
                CT_Rst i1 = stbl.GetEntryAt(i);
                String i2 = rtbl.GetEntryAt(i);
                ClassicAssert.AreEqual(i1.t, i2);
            }

        }

        //51519
        [Test]
        public void TestPhoneticRuns()
        {

            OPCPackage pkg = OPCPackage.Open(_ssTests.OpenResourceAsStream("51519.xlsx"));
            List<PackagePart> parts = pkg.GetPartsByName(new Regex("/xl/sharedStrings.xml", RegexOptions.Compiled));
            ClassicAssert.AreEqual(1, parts.Count);

            ReadOnlySharedStringsTable rtbl = new ReadOnlySharedStringsTable(parts[0], true);
            List<String> strings = rtbl.Items;
            ClassicAssert.AreEqual(49, strings.Count);

            ClassicAssert.AreEqual("\u30B3\u30E1\u30F3\u30C8", rtbl.GetEntryAt(0));
            ClassicAssert.AreEqual("\u65E5\u672C\u30AA\u30E9\u30AF\u30EB \u30CB\u30DB\u30F3", rtbl.GetEntryAt(3));

            //now do not include phonetic runs
            rtbl = new ReadOnlySharedStringsTable(parts[0], false);
            strings = rtbl.Items;
            ClassicAssert.AreEqual(49, strings.Count);

            ClassicAssert.AreEqual("\u30B3\u30E1\u30F3\u30C8", rtbl.GetEntryAt(0));
            ClassicAssert.AreEqual("\u65E5\u672C\u30AA\u30E9\u30AF\u30EB", rtbl.GetEntryAt(3));

        }
        [Test]
        public void TestEmptySSTOnPackageObtainedViaWorkbook()
        {

            XSSFWorkbook wb = new XSSFWorkbook(_ssTests.OpenResourceAsStream("noSharedStringTable.xlsx"));
            OPCPackage pkg = wb.Package;
            assertEmptySST(pkg);
            wb.Close();
        }
        [Test]
        public void TestEmptySSTOnPackageDirect()
        {

            OPCPackage pkg = OPCPackage.Open(_ssTests.OpenResourceAsStream("noSharedStringTable.xlsx"));
            assertEmptySST(pkg);
        }

        private void assertEmptySST(OPCPackage pkg)
        {

            ReadOnlySharedStringsTable sst = new ReadOnlySharedStringsTable(pkg);
            ClassicAssert.AreEqual(0, sst.Count);
            ClassicAssert.AreEqual(0, sst.UniqueCount);
            ClassicAssert.IsNull(sst.Items); // same state it's left in if fed a package which has no SST part.
        }

    }
}

