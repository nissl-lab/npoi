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

// Regression tests for GitHub issue #1844:
// WorkbookFactory.Create(Stream) hangs in an infinite loop on malformed .xls
// whose document-property Size field claims more bytes than the actual block chain holds.

using System;
using System.IO;
using NUnit.Framework;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOI.Util;

namespace TestCases.POIFS.FileSystem
{
    [TestFixture]
    public class TestBug1844
    {
        // OLE2 property-entry layout constants (ECMA-376 / OLE specification):
        // Each property entry is 128 bytes.  The Size field is a little-endian
        // int32 at offset 0x78 within the entry.
        private const int PropertyEntrySize = 128;
        private const int PropertySizeOffset = 0x78; // within a single property entry

        /// <summary>
        /// Directly tests that NDocumentInputStream.ReadFully throws IOException
        /// (not hang forever) when the block chain ends before _document_size bytes
        /// have been consumed.  This is the root fix for issue #1844.
        ///
        /// Strategy: write a 1-block document (512 bytes of data), then attempt to
        /// read 1024 bytes from it.  The iterator exhausts after the first block and
        /// MoveNext() returns false; our fix must throw IOException rather than loop.
        /// </summary>
        [Test]
        public void ReadFullyThrowsIOExceptionWhenChainExhausted()
        {
            // Write exactly 4097 bytes so the document uses big-block storage
            // (threshold is 4096 bytes).  The real chain spans 9 sectors (ceil(4097/512)).
            byte[] bigPayload = new byte[4097];
            for (int i = 0; i < bigPayload.Length; i++) bigPayload[i] = (byte)(i & 0xFF);

            var pfs = new NPOIFSFileSystem();
            pfs.Root.CreateDocument("BigDoc", new MemoryStream(bigPayload));
            DocumentEntry bigEntry = (DocumentEntry)pfs.Root.GetEntry("BigDoc");
            var bigDoc = new NPOIFSDocument((DocumentNode)bigEntry);

            // Inflate the property Size so CheckAvaliable passes the initial guard,
            // simulating a malformed OLE2 file where the property claims far more
            // bytes than the actual block chain holds.
            bigDoc.DocumentProperty.Size = int.MaxValue;

            var dis = new NDocumentInputStream(bigDoc);
            byte[] buf = new byte[999_999];

            // Before the fix this would spin forever; after the fix it must throw.
            Assert.Throws<IOException>(() => dis.ReadFully(buf, 0, buf.Length),
                "NDocumentInputStream.ReadFully must throw IOException when the block " +
                "chain is shorter than the requested read length.");
        }

        /// <summary>
        /// High-level regression: opening an OLE2 workbook whose Workbook document-
        /// property Size is inflated to int.MaxValue must terminate with an exception
        /// rather than hanging.  Before the fix the call would loop forever.
        /// </summary>
        [Test]
        public void InflatedDocumentSizeThrowsException()
        {
            byte[] rawFile = BuildMinimalHssfBytes();

            // Patch the "Workbook" document-property Size to int.MaxValue.
            // The property table starts right after the header (sector 0 → file
            // offset 512 for 512-byte sectors). Property #0 is the root entry (128 B);
            // property #1 is the "Workbook" document entry.
            int propertyTableFileOffset = 512;
            int workbookPropertyOffset  = propertyTableFileOffset + PropertyEntrySize;
            int sizeFieldOffset         = workbookPropertyOffset + PropertySizeOffset;
            LittleEndian.PutInt(rawFile, sizeFieldOffset, int.MaxValue);

            using var ms = new MemoryStream(rawFile, writable: false);
            // Any exception is acceptable — the important invariant is that the
            // call terminates instead of hanging.
            Assert.Catch<Exception>(() =>
            {
                var pfs = new POIFSFileSystem(ms);
                using var wb = new HSSFWorkbook(pfs.Root, true);
            },
            "Opening a workbook with an inflated document-property Size must throw " +
            "rather than hang indefinitely.");
        }

        // ------------------------------------------------------------------
        // Helpers
        // ------------------------------------------------------------------

        /// <summary>
        /// Creates a minimal valid .xls file in memory and returns its raw bytes.
        /// </summary>
        private static byte[] BuildMinimalHssfBytes()
        {
            var wb = new HSSFWorkbook();
            wb.CreateSheet("Sheet1");
            using var ms = new MemoryStream();
            wb.Write(ms);
            wb.Close();
            return ms.ToArray();
        }
    }
}
