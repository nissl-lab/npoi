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

using NPOI.HWPF.Model.IO;
using NUnit.Framework;
using TestCases.HWPF;
namespace NPOI.HWPF.Model
{
    [TestFixture]
    public class TestStyleSheet
    {
        private StyleSheet _styleSheet = null;
        private HWPFDocFixture _hWPFDocFixture;
        [Test]
        public void TestReadWrite()
        {
            HWPFFileSystem fileSys = new HWPFFileSystem();


            HWPFStream tableOut = fileSys.GetStream("1Table");
            HWPFStream mainOut = fileSys.GetStream("WordDocument");

            _styleSheet.WriteTo(tableOut);

            byte[] newTableStream = tableOut.ToArray();

            StyleSheet newStyleSheet = new StyleSheet(newTableStream, 0);
            Assert.AreEqual(newStyleSheet, _styleSheet);

        }
        [Test]
        public void TestReadWriteFromNonZeroOffset()
        {
            HWPFFileSystem fileSys = new HWPFFileSystem();
            HWPFStream tableOut = fileSys.GetStream("1Table");

            tableOut.Write(new byte[20]); // 20 bytes of whatever at the front.
            _styleSheet.WriteTo(tableOut);

            byte[] newTableStream = tableOut.ToArray();

            StyleSheet newStyleSheet = new StyleSheet(newTableStream, 20);
            Assert.AreEqual(newStyleSheet, _styleSheet);
        }
        [SetUp]
        public void SetUp()
        {
            /**@todo verify the constructors*/
            _hWPFDocFixture = new HWPFDocFixture(this);
            _hWPFDocFixture.SetUp();
            FileInformationBlock fib = _hWPFDocFixture._fib;
            byte[] mainStream = _hWPFDocFixture._mainStream;
            byte[] tableStream = _hWPFDocFixture._tableStream;

            _hWPFDocFixture.SetUp();
            _styleSheet = new StyleSheet(tableStream, fib.GetFcStshf());
        }
        [TearDown]
        public void TearDown()
        {
            _styleSheet = null;

            _hWPFDocFixture = null;
        }

    }
}
