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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.POIFS.FileSystem;
using NUnit.Framework;using NUnit.Framework.Legacy;
using NPOI.Util;
using System.IO;

namespace TestCases.POIFS.FileSystem
{
    [TestFixture]
    public class TestEntryUtils
    {
        private byte[] dataSmallA = new byte[] { 12, 42, 11, unchecked((byte)-12), unchecked((byte)-121) };
        private byte[] dataSmallB = new byte[] { 11, 73, 21, unchecked((byte)-92), unchecked((byte)-103) };
        [Test]
        public void TestCopyRecursively()
        {
            POIFSFileSystem fsD = new POIFSFileSystem();
            POIFSFileSystem fs = new POIFSFileSystem();
            DirectoryEntry dirA = fs.CreateDirectory("DirA");
            DirectoryEntry dirB = fs.CreateDirectory("DirB");

            DocumentEntry entryR = fs.CreateDocument(new ByteArrayInputStream(dataSmallA), "EntryRoot");
            DocumentEntry entryA1 = dirA.CreateDocument("EntryA1", new ByteArrayInputStream(dataSmallA));
            DocumentEntry entryA2 = dirA.CreateDocument("EntryA2", new ByteArrayInputStream(dataSmallB));

            // Copy docs
            ClassicAssert.AreEqual(0, fsD.Root.EntryCount);
            EntryUtils.CopyNodeRecursively(entryR, fsD.Root);

            ClassicAssert.AreEqual(1, fsD.Root.EntryCount);
            ClassicAssert.IsNotNull(fsD.Root.GetEntry("EntryRoot"));

            EntryUtils.CopyNodeRecursively(entryA1, fsD.Root);
            ClassicAssert.AreEqual(2, fsD.Root.EntryCount);
            ClassicAssert.IsNotNull(fsD.Root.GetEntry("EntryRoot"));
            ClassicAssert.IsNotNull(fsD.Root.GetEntry("EntryA1"));

            EntryUtils.CopyNodeRecursively(entryA2, fsD.Root);
            ClassicAssert.AreEqual(3, fsD.Root.EntryCount);
            ClassicAssert.IsNotNull(fsD.Root.GetEntry("EntryRoot"));
            ClassicAssert.IsNotNull(fsD.Root.GetEntry("EntryA1"));
            ClassicAssert.IsNotNull(fsD.Root.GetEntry("EntryA2"));

            // Copy directories
            fsD = new POIFSFileSystem();
            ClassicAssert.AreEqual(0, fsD.Root.EntryCount);

            EntryUtils.CopyNodeRecursively(dirB, fsD.Root);
            ClassicAssert.AreEqual(1, fsD.Root.EntryCount);
            ClassicAssert.IsNotNull(fsD.Root.GetEntry("DirB"));
            ClassicAssert.AreEqual(0, ((DirectoryEntry)fsD.Root.GetEntry("DirB")).EntryCount);

            EntryUtils.CopyNodeRecursively(dirA, fsD.Root);
            ClassicAssert.AreEqual(2, fsD.Root.EntryCount);
            ClassicAssert.IsNotNull(fsD.Root.GetEntry("DirB"));
            ClassicAssert.AreEqual(0, ((DirectoryEntry)fsD.Root.GetEntry("DirB")).EntryCount);
            ClassicAssert.IsNotNull(fsD.Root.GetEntry("DirA"));
            ClassicAssert.AreEqual(2, ((DirectoryEntry)fsD.Root.GetEntry("DirA")).EntryCount);

            // Copy the whole lot
            fsD = new POIFSFileSystem();
            ClassicAssert.AreEqual(0, fsD.Root.EntryCount);

            EntryUtils.CopyNodes(fs, fsD, new List<String>());
            ClassicAssert.AreEqual(3, fsD.Root.EntryCount);
            ClassicAssert.IsNotNull(fsD.Root.GetEntry(dirA.Name));
            ClassicAssert.IsNotNull(fsD.Root.GetEntry(dirB.Name));
            ClassicAssert.IsNotNull(fsD.Root.GetEntry(entryR.Name));
            ClassicAssert.AreEqual(0, ((DirectoryEntry)fsD.Root.GetEntry("DirB")).EntryCount);
            ClassicAssert.AreEqual(2, ((DirectoryEntry)fsD.Root.GetEntry("DirA")).EntryCount);
        }
        [Test]
        public void TestAreDocumentsIdentical()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            DirectoryEntry dirA = fs.CreateDirectory("DirA");
            DirectoryEntry dirB = fs.CreateDirectory("DirB");

            DocumentEntry entryA1 = dirA.CreateDocument("Entry1", new ByteArrayInputStream(dataSmallA));
            DocumentEntry entryA1b = dirA.CreateDocument("Entry1b", new ByteArrayInputStream(dataSmallA));
            DocumentEntry entryA2 = dirA.CreateDocument("Entry2", new ByteArrayInputStream(dataSmallB));
            DocumentEntry entryB1 = dirB.CreateDocument("Entry1", new ByteArrayInputStream(dataSmallA));


            // Names must match
            ClassicAssert.AreEqual(false, entryA1.Name.Equals(entryA1b.Name));
            ClassicAssert.AreEqual(false, EntryUtils.AreDocumentsIdentical(entryA1, entryA1b));

            // Contents must match
            ClassicAssert.AreEqual(false, EntryUtils.AreDocumentsIdentical(entryA1, entryA2));

            // Parents don't matter if contents + names are the same
            ClassicAssert.AreEqual(false, entryA1.Parent.Equals(entryB1.Parent));
            ClassicAssert.AreEqual(true, EntryUtils.AreDocumentsIdentical(entryA1, entryB1));


            // Can work with NPOIFS + POIFS
            //ByteArrayOutputStream tmpO = new ByteArrayOutputStream();
            MemoryStream tmpO = new MemoryStream();
            fs.WriteFileSystem(tmpO);
            ByteArrayInputStream tmpI = new ByteArrayInputStream(tmpO.ToArray());
            NPOIFSFileSystem nfs = new NPOIFSFileSystem(tmpI);

            DirectoryEntry dN1 = (DirectoryEntry)nfs.Root.GetEntry("DirA");
            DirectoryEntry dN2 = (DirectoryEntry)nfs.Root.GetEntry("DirB");
            DocumentEntry eNA1 = (DocumentEntry)dN1.GetEntry(entryA1.Name);
            DocumentEntry eNA2 = (DocumentEntry)dN1.GetEntry(entryA2.Name);
            DocumentEntry eNB1 = (DocumentEntry)dN2.GetEntry(entryB1.Name);

            ClassicAssert.AreEqual(false, EntryUtils.AreDocumentsIdentical(eNA1, eNA2));
            ClassicAssert.AreEqual(true, EntryUtils.AreDocumentsIdentical(eNA1, eNB1));

            ClassicAssert.AreEqual(false, EntryUtils.AreDocumentsIdentical(eNA1, entryA1b));
            ClassicAssert.AreEqual(false, EntryUtils.AreDocumentsIdentical(eNA1, entryA2));

            ClassicAssert.AreEqual(true, EntryUtils.AreDocumentsIdentical(eNA1, entryA1));
            ClassicAssert.AreEqual(true, EntryUtils.AreDocumentsIdentical(eNA1, entryB1));
        }
        [Test]
        public void TestAreDirectoriesIdentical()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            DirectoryEntry dirA = fs.CreateDirectory("DirA");
            DirectoryEntry dirB = fs.CreateDirectory("DirB");

            // Names must match
            ClassicAssert.AreEqual(false, EntryUtils.AreDirectoriesIdentical(dirA, dirB));

            // Empty dirs are fine
            DirectoryEntry dirA1 = dirA.CreateDirectory("TheDir");
            DirectoryEntry dirB1 = dirB.CreateDirectory("TheDir");
            ClassicAssert.AreEqual(0, dirA1.EntryCount);
            ClassicAssert.AreEqual(0, dirB1.EntryCount);
            ClassicAssert.AreEqual(true, EntryUtils.AreDirectoriesIdentical(dirA1, dirB1));

            // Otherwise children must match
            dirA1.CreateDocument("Entry1", new ByteArrayInputStream(dataSmallA));
            ClassicAssert.AreEqual(false, EntryUtils.AreDirectoriesIdentical(dirA1, dirB1));

            dirB1.CreateDocument("Entry1", new ByteArrayInputStream(dataSmallA));
            ClassicAssert.AreEqual(true, EntryUtils.AreDirectoriesIdentical(dirA1, dirB1));

            dirA1.CreateDirectory("DD");
            ClassicAssert.AreEqual(false, EntryUtils.AreDirectoriesIdentical(dirA1, dirB1));
            dirB1.CreateDirectory("DD");
            ClassicAssert.AreEqual(true, EntryUtils.AreDirectoriesIdentical(dirA1, dirB1));


            // Excludes support
            List<String> excl = new List<string>(new String[] { "Ignore1", "IgnDir/Ign2" });
            FilteringDirectoryNode fdA = new FilteringDirectoryNode(dirA1, excl);
            FilteringDirectoryNode fdB = new FilteringDirectoryNode(dirB1, excl);

            ClassicAssert.AreEqual(true, EntryUtils.AreDirectoriesIdentical(fdA, fdB));

            // Add an ignored doc, no notice is taken
            fdA.CreateDocument("Ignore1", new ByteArrayInputStream(dataSmallA));
            ClassicAssert.AreEqual(true, EntryUtils.AreDirectoriesIdentical(fdA, fdB));

            // Add a directory with filtered contents, not the same
            DirectoryEntry dirAI = dirA1.CreateDirectory("IgnDir");
            ClassicAssert.AreEqual(false, EntryUtils.AreDirectoriesIdentical(fdA, fdB));

            DirectoryEntry dirBI = dirB1.CreateDirectory("IgnDir");
            ClassicAssert.AreEqual(true, EntryUtils.AreDirectoriesIdentical(fdA, fdB));

            // Add something to the filtered subdir that gets ignored
            dirAI.CreateDocument("Ign2", new ByteArrayInputStream(dataSmallA));
            ClassicAssert.AreEqual(true, EntryUtils.AreDirectoriesIdentical(fdA, fdB));

            // And something that doesn't
            dirAI.CreateDocument("IgnZZ", new ByteArrayInputStream(dataSmallA));
            ClassicAssert.AreEqual(false, EntryUtils.AreDirectoriesIdentical(fdA, fdB));

            dirBI.CreateDocument("IgnZZ", new ByteArrayInputStream(dataSmallA));
            ClassicAssert.AreEqual(true, EntryUtils.AreDirectoriesIdentical(fdA, fdB));
        }
    }

}
