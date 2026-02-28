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
using TestCases.HWPF;
using System.Collections.Generic;
using NUnit.Framework;
namespace NPOI.HWPF.Model
{
    class TextPieceTable1 : TextPieceTable
    {
        public override bool IsIndexInTable(int bytePos)
        {
            return true;
        }
    }
    [TestFixture]
    public class TestCHPBinTable
    {
        private CHPBinTable _cHPBinTable = null;
        private HWPFDocFixture _hWPFDocFixture;

        private TextPieceTable fakeTPT = new TextPieceTable1();

        [Test]
        public void TestReadWrite()
        {
            FileInformationBlock fib = _hWPFDocFixture._fib;
            byte[] mainStream = _hWPFDocFixture._mainStream;
            byte[] tableStream = _hWPFDocFixture._tableStream;
            int fcMin = fib.GetFcMin();

            _cHPBinTable = new CHPBinTable(mainStream, tableStream, fib.GetFcPlcfbteChpx(), fib.GetLcbPlcfbteChpx(), fcMin, fakeTPT);

            HWPFFileSystem fileSys = new HWPFFileSystem();

            _cHPBinTable.WriteTo(fileSys, 0);
            MemoryStream tableOut = fileSys.GetStream("1Table");
            MemoryStream mainOut = fileSys.GetStream("WordDocument");

            byte[] newTableStream = tableOut.ToArray();
            byte[] newMainStream = mainOut.ToArray();

            CHPBinTable newBinTable = new CHPBinTable(newMainStream, newTableStream, 0, newTableStream.Length, 0, fakeTPT);

            List<CHPX> oldTextRuns = _cHPBinTable._textRuns;
            List<CHPX> newTextRuns = newBinTable._textRuns;

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

            _hWPFDocFixture = new HWPFDocFixture(this);
            _hWPFDocFixture.SetUp();
        }
        [TearDown]
        public void tearDown()
        {
            _cHPBinTable = null;
            _hWPFDocFixture = null;
        }

    }

}