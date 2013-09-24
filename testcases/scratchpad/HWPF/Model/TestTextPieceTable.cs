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
using System.IO;
using TestCases.HWPF;
namespace NPOI.HWPF.Model
{


    [TestFixture]
    public class TestTextPieceTable
    {
        private HWPFDocFixture _hWPFDocFixture;
        //private String dirname;
        [Test]
        public void TestReadWrite()
        {
            FileInformationBlock fib = _hWPFDocFixture._fib;
            byte[] mainStream = _hWPFDocFixture._mainStream;
            byte[] tableStream = _hWPFDocFixture._tableStream;
            int fcMin = fib.GetFcMin();

            ComplexFileTable cft = new ComplexFileTable(mainStream, tableStream, fib.GetFcClx(), fcMin);


            HWPFFileSystem fileSys = new HWPFFileSystem();

            cft.WriteTo(fileSys);
            MemoryStream tableOut = fileSys.GetStream("1Table");
            MemoryStream mainOut = fileSys.GetStream("WordDocument");

            byte[] newTableStream = tableOut.ToArray();
            byte[] newMainStream = mainOut.ToArray();

            ComplexFileTable newCft = new ComplexFileTable(newMainStream, newTableStream, 0, 0);

            TextPieceTable oldTextPieceTable = cft.GetTextPieceTable();
            TextPieceTable newTextPieceTable = newCft.GetTextPieceTable();

            Assert.AreEqual(oldTextPieceTable.Text.ToString(), newTextPieceTable.Text.ToString());
        }

        /**
         * Check that we do the positions correctly when
         *  working with pure-ascii
         */
        [Test]
        public void TestAsciiParts()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("ThreeColHeadFoot.doc");
            TextPieceTable tbl = doc.TextTable;

            // All ascii, so stored in one big lump
            Assert.AreEqual(1, tbl.TextPieces.Count);
            TextPiece tp = (TextPiece)tbl.TextPieces[0];

            Assert.AreEqual(0, tp.Start);
            Assert.AreEqual(339, tp.End);
            Assert.AreEqual(339, tp.CharacterLength);
            Assert.AreEqual(339, tp.BytesLength);
            Assert.IsTrue(tp.GetStringBuilder().ToString().StartsWith("This is a sample word document"));


            // Save and re-load
            HWPFDocument docB = SaveAndReload(doc);
            tbl = docB.TextTable;

            Assert.AreEqual(1, tbl.TextPieces.Count);
            tp = (TextPiece)tbl.TextPieces[0];

            Assert.AreEqual(0, tp.Start);
            Assert.AreEqual(339, tp.End);
            Assert.AreEqual(339, tp.CharacterLength);
            Assert.AreEqual(339, tp.BytesLength);
            Assert.IsTrue(tp.GetStringBuilder().ToString().StartsWith("This is a sample word document"));
        }

        /**
         * Check that we do the positions correctly when
         *  working with a mix ascii, unicode file
         */
        [Test]
        public void TestUnicodeParts()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("HeaderFooterUnicode.doc");
            TextPieceTable tbl = doc.TextTable;

            // In three bits, split every 512 bytes
            Assert.AreEqual(3, tbl.TextPieces.Count);
            TextPiece tpA = (TextPiece)tbl.TextPieces[0];
            TextPiece tpB = (TextPiece)tbl.TextPieces[1];
            TextPiece tpC = (TextPiece)tbl.TextPieces[2];

            Assert.IsTrue(tpA.IsUnicode);
            Assert.IsTrue(tpB.IsUnicode);
            Assert.IsTrue(tpC.IsUnicode);

            Assert.AreEqual(256, tpA.CharacterLength);
            Assert.AreEqual(256, tpB.CharacterLength);
            Assert.AreEqual(19, tpC.CharacterLength);

            Assert.AreEqual(512, tpA.BytesLength);
            Assert.AreEqual(512, tpB.BytesLength);
            Assert.AreEqual(38, tpC.BytesLength);

            Assert.AreEqual(0, tpA.Start);
            Assert.AreEqual(256, tpA.End);
            Assert.AreEqual(256, tpB.Start);
            Assert.AreEqual(512, tpB.End);
            Assert.AreEqual(512, tpC.Start);
            Assert.AreEqual(531, tpC.End);


            // Save and re-load
            HWPFDocument docB = SaveAndReload(doc);
            tbl = docB.TextTable;

            Assert.AreEqual(3, tbl.TextPieces.Count);
            tpA = (TextPiece)tbl.TextPieces[0];
            tpB = (TextPiece)tbl.TextPieces[1];
            tpC = (TextPiece)tbl.TextPieces[2];

            Assert.IsTrue(tpA.IsUnicode);
            Assert.IsTrue(tpB.IsUnicode);
            Assert.IsTrue(tpC.IsUnicode);

            Assert.AreEqual(256, tpA.CharacterLength);
            Assert.AreEqual(256, tpB.CharacterLength);
            Assert.AreEqual(19, tpC.CharacterLength);

            Assert.AreEqual(512, tpA.BytesLength);
            Assert.AreEqual(512, tpB.BytesLength);
            Assert.AreEqual(38, tpC.BytesLength);

            Assert.AreEqual(0, tpA.Start);
            Assert.AreEqual(256, tpA.End);
            Assert.AreEqual(256, tpB.Start);
            Assert.AreEqual(512, tpB.End);
            Assert.AreEqual(512, tpC.Start);
            Assert.AreEqual(531, tpC.End);
        }

        protected HWPFDocument SaveAndReload(HWPFDocument doc)
        {
            MemoryStream baos = new MemoryStream();
            doc.Write(baos);

            return new HWPFDocument(
                    new MemoryStream(baos.ToArray())
            );
        }
        [SetUp]
        public void SetUp()
        {
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

