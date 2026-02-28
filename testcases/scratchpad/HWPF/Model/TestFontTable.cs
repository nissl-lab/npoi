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


namespace TestCases.HWPF.Model
{
    
    using NPOI.HWPF.Model.IO;
    using NPOI.HWPF.Model;
    using NUnit.Framework;

    [TestFixture]
    public class TestFontTable
    {
        private FontTable _fontTable = null;
        private HWPFDocFixture _hWPFDocFixture;
        [Test]
        public void TestReadWrite()
        {
            FileInformationBlock fib = _hWPFDocFixture._fib;
            byte[] tableStream = _hWPFDocFixture._tableStream;

            int fcSttbfffn = fib.GetFcSttbfffn();
            int lcbSttbfffn = fib.GetLcbSttbfffn();

            _fontTable = new FontTable(tableStream, fcSttbfffn, lcbSttbfffn);

            HWPFFileSystem fileSys = new HWPFFileSystem();

            _fontTable.WriteTo(fileSys);
            HWPFStream tableOut = fileSys.GetStream("1Table");


            byte[] newTableStream = tableOut.ToArray();


            FontTable newFontTable = new FontTable(newTableStream, 0, newTableStream.Length);

            Assert.IsTrue(_fontTable.Equals(newFontTable));

        }
        [SetUp]
        public void SetUp()
        {
            /**@todo verify the constructors*/
            _hWPFDocFixture = new HWPFDocFixture(this);

            _hWPFDocFixture.SetUp();
        }
        [TearDown]
        public void TearDown()
        {
            _hWPFDocFixture = null;
        }

    }
}


