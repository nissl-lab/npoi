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
using System.IO;
using System.Collections.Generic;

using TestCases.HWPF;
using NUnit.Framework;
namespace NPOI.HWPF.Model
{
    [TestFixture]
    public class TestPAPBinTable
    {
        private PAPBinTable _pAPBinTable = null;
        private HWPFDocFixture _hWPFDocFixture;

        [Test]
        public void TestReadWrite()
        {
            TextPieceTable fakeTPT = new TextPieceTable();

            FileInformationBlock fib = _hWPFDocFixture._fib;
            byte[] mainStream = _hWPFDocFixture._mainStream;
            byte[] tableStream = _hWPFDocFixture._tableStream;

            _pAPBinTable = new PAPBinTable(mainStream, tableStream, null, fib.GetFcPlcfbtePapx(), fib.GetLcbPlcfbtePapx(), fakeTPT);

            HWPFFileSystem fileSys = new HWPFFileSystem();

            _pAPBinTable.WriteTo(fileSys, fakeTPT);
            MemoryStream tableOut = fileSys.GetStream("1Table");
            MemoryStream mainOut = fileSys.GetStream("WordDocument");

            byte[] newTableStream = tableOut.ToArray();
            byte[] newMainStream = mainOut.ToArray();

            PAPBinTable newBinTable = new PAPBinTable(newMainStream, newTableStream, null, 0, newTableStream.Length, fakeTPT);

            List<PAPX> oldTextRuns = _pAPBinTable.GetParagraphs();
            List<PAPX> newTextRuns = newBinTable.GetParagraphs();

            Assert.AreEqual(oldTextRuns.Count, newTextRuns.Count);

            int size = oldTextRuns.Count;
            for (int x = 0; x < size; x++)
            {
                PropertyNode oldNode = (PropertyNode)oldTextRuns[x];
                PropertyNode newNode = (PropertyNode)newTextRuns[x];

                Assert.IsTrue(oldNode.Equals(newNode));
            }




        }
        [SetUp]
        public void SetUp()
        {
            /**@todo verify the constructors*/
            _hWPFDocFixture = new HWPFDocFixture(this);

            _hWPFDocFixture.SetUp();
        }
        [TearDown]
        public void tearDown()
        {
            _hWPFDocFixture = null;
        }

    }
}
