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
    using NPOI.HWPF.Model;
    using NPOI.HWPF.Model.IO;
    using System.IO;
    
    using System.Collections.Generic;
    using NUnit.Framework;

    [TestFixture]
    public class TestSectionTable
    {
        private HWPFDocFixture _hWPFDocFixture;
        [Test]
        public void TestReadWrite()
        {
            FileInformationBlock fib = _hWPFDocFixture._fib;
            byte[] mainStream = _hWPFDocFixture._mainStream;
            byte[] tableStream = _hWPFDocFixture._tableStream;
            int fcMin = fib.GetFcMin();


            ComplexFileTable cft = new ComplexFileTable(mainStream, tableStream, fib.GetFcClx(), fcMin);
            TextPieceTable tpt = cft.GetTextPieceTable();

            SectionTable sectionTable = new SectionTable(mainStream, tableStream,
                                                         fib.GetFcPlcfsed(),
                                                         fib.GetLcbPlcfsed(),
                                                         fcMin, tpt, fib.GetSubdocumentTextStreamLength(SubdocumentType.MAIN));
            HWPFFileSystem fileSys = new HWPFFileSystem();

            sectionTable.WriteTo(fileSys, 0);
            MemoryStream tableOut = fileSys.GetStream("1Table");
            MemoryStream mainOut = fileSys.GetStream("WordDocument");

            byte[] newTableStream = tableOut.ToArray();
            byte[] newMainStream = mainOut.ToArray();

            SectionTable newSectionTable = new SectionTable(
                    newMainStream, newTableStream, 0,
                    newTableStream.Length, 0, tpt, fib.GetSubdocumentTextStreamLength(SubdocumentType.MAIN));

            List<SEPX> oldSections = sectionTable.GetSections();
            List<SEPX> newSections = newSectionTable.GetSections();

            Assert.AreEqual(oldSections.Count, newSections.Count);

            //test for proper char offset conversions
            PlexOfCps oldSedPlex = new PlexOfCps(tableStream, fib.GetFcPlcfsed(),
                                                              fib.GetLcbPlcfsed(), 12);
            PlexOfCps newSedPlex = new PlexOfCps(newTableStream, 0,
                                                 newTableStream.Length, 12);
            Assert.AreEqual(oldSedPlex.Length, newSedPlex.Length);

            for (int x = 0; x < oldSedPlex.Length; x++)
            {
                Assert.AreEqual(oldSedPlex.GetProperty(x).Start, newSedPlex.GetProperty(x).Start);
                Assert.AreEqual(oldSedPlex.GetProperty(x).End, newSedPlex.GetProperty(x).End);
            }

            int size = oldSections.Count;
            for (int x = 0; x < size; x++)
            {
                PropertyNode oldNode = (PropertyNode)oldSections[x];
                PropertyNode newNode = (PropertyNode)newSections[x];
                Assert.AreEqual(oldNode, newNode);
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
        public void TearDown()
        {
            _hWPFDocFixture = null;
        }

    }

}