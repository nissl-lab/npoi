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

namespace TestCases.OpenXml4Net.OPC
{
    using System;
    using System.IO;
    using SysZip = System.IO.Compression;
    using ICSharpCode.SharpZipLib.Zip;
    using NPOI.OpenXml4Net.Util;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    /// <summary>
    /// Regression tests for the OOXML decompression-bomb (zip-bomb) denial-of-service
    /// fix (GHSA-6rjh-63f8-8p48). A crafted package used to allocate memory proportional
    /// to the uncompressed content with no effective limit; the read paths now enforce an
    /// absolute per-entry size cap and a minimum inflate ratio via ZipSecureFile.
    /// </summary>
    [TestFixture]
    public class TestZipBomb
    {
        // ---- ZipSecureFile.CheckThreshold: the pure limit logic -------------------------

        [Test]
        public void CheckThresholdAllowsSmallExpandedSize()
        {
            // Below the 100 KB grace size, any ratio is tolerated (avoids false positives).
            Assert.DoesNotThrow(() => ZipSecureFile.CheckThreshold(50 * 1024, 1));
        }

        [Test]
        public void CheckThresholdAllowsHealthyRatio()
        {
            // Above grace size but with a reasonable compression ratio (0.5 >= 0.01).
            Assert.DoesNotThrow(() => ZipSecureFile.CheckThreshold(1_000_000, 500_000));
        }

        [Test]
        public void CheckThresholdRejectsBombRatio()
        {
            // Above grace size, ratio far below the 1% minimum -> zip bomb.
            IOException ex = Assert.Throws<IOException>(
                () => ZipSecureFile.CheckThreshold(200L * 1024 * 1024, 1024));
            Assert.That(ex.Message, Does.Contain("Zip bomb"));
        }

        [Test]
        public void CheckThresholdWithUnknownCompressedSizeUsesAbsoluteCap()
        {
            // Unknown compressed size (streamed / data-descriptor entry): only the absolute
            // cap applies, so a sub-cap size passes but exceeding MAX_ENTRY_SIZE throws.
            Assert.DoesNotThrow(() => ZipSecureFile.CheckThreshold(200L * 1024 * 1024, -1));
            IOException ex = Assert.Throws<IOException>(
                () => ZipSecureFile.CheckThreshold(5L * 1024 * 1024 * 1024, -1));
            Assert.That(ex.Message, Does.Contain("Zip bomb"));
        }

        // ---- Seekable read path (ZipFileZipEntrySource) --------------------------------

        [Test]
        public void SeekableZipBombEntryIsRejected()
        {
            byte[] zipBytes = BuildRepeatedByteZip("bomb.bin", 16L * 1024 * 1024, dataDescriptor: false);
            using MemoryStream ms = new MemoryStream(zipBytes);
            ZipFile zf = new ZipFile(ms);
            ZipFileZipEntrySource src = new ZipFileZipEntrySource(zf);
            ZipEntry entry = FirstEntry(src);

            using Stream guarded = src.GetInputStream(entry);
            IOException ex = Assert.Throws<IOException>(() => Drain(guarded));
            Assert.That(ex.Message, Does.Contain("Zip bomb"));
            src.Close();
        }

        [Test]
        public void LegitimateSeekableEntryIsReadInFull()
        {
            // ~256 KB of incompressible data: ratio ~1.0 and above the grace size, so it
            // must be read in full without a false-positive zip-bomb exception.
            byte[] payload = new byte[256 * 1024];
            new Random(12345).NextBytes(payload);
            byte[] zipBytes = BuildRawZip("data.bin", payload, dataDescriptor: false);

            using MemoryStream ms = new MemoryStream(zipBytes);
            ZipFile zf = new ZipFile(ms);
            ZipFileZipEntrySource src = new ZipFileZipEntrySource(zf);
            ZipEntry entry = FirstEntry(src);

            using Stream guarded = src.GetInputStream(entry);
            long total = Drain(guarded);
            ClassicAssert.AreEqual(payload.Length, total, "legitimate entry must be readable in full");
            src.Close();
        }

        // ---- Streamed read path (ZipInputStreamZipEntrySource + CountingStream) ---------

        [Test]
        public void DataDescriptorZipBombEntryIsRejected()
        {
            // Data-descriptor entries report Size == -1, which historically bypassed the
            // only size guard. The CountingStream-backed ratio check (wired exactly as
            // ZipPackage(Stream) wires it) must still reject the bomb.
            byte[] zipBytes = BuildRepeatedByteZip("bomb.bin", 16L * 1024 * 1024, dataDescriptor: true);
            using MemoryStream ms = new MemoryStream(zipBytes);
            CountingStream counter = new CountingStream(ms);
            using ZipInputStream zis = new ZipInputStream(counter);

            IOException ex = Assert.Throws<IOException>(
                () => new ZipInputStreamZipEntrySource(zis, counter));
            Assert.That(ex.Message, Does.Contain("Zip bomb"));
        }

        // ---- helpers -------------------------------------------------------------------

        private static ZipEntry FirstEntry(ZipEntrySource src)
        {
            var entries = src.Entries;
            ClassicAssert.IsTrue(entries.MoveNext(), "zip should contain at least one entry");
            return (ZipEntry)entries.Current;
        }

        private static long Drain(Stream s)
        {
            byte[] buffer = new byte[8192];
            long total = 0;
            int read;
            while ((read = s.Read(buffer, 0, buffer.Length)) > 0)
            {
                total += read;
            }
            return total;
        }

        /// <summary>Builds a zip whose single entry is <paramref name="uncompressedBytes"/> copies of 'A'.</summary>
        private static byte[] BuildRepeatedByteZip(string entryName, long uncompressedBytes, bool dataDescriptor)
        {
            return BuildZip(entryName, dataDescriptor, entryStream =>
            {
                byte[] chunk = new byte[64 * 1024];
                for (int i = 0; i < chunk.Length; i++)
                {
                    chunk[i] = (byte)'A';
                }
                long remaining = uncompressedBytes;
                while (remaining > 0)
                {
                    int n = (int)Math.Min(chunk.Length, remaining);
                    entryStream.Write(chunk, 0, n);
                    remaining -= n;
                }
            });
        }

        private static byte[] BuildRawZip(string entryName, byte[] payload, bool dataDescriptor)
        {
            return BuildZip(entryName, dataDescriptor, entryStream => entryStream.Write(payload, 0, payload.Length));
        }

        private static byte[] BuildZip(string entryName, bool dataDescriptor, Action<Stream> writeEntry)
        {
            MemoryStream backing = new MemoryStream();
            // A non-seekable wrapper forces System.IO.Compression to emit data descriptors
            // (general-purpose bit 3), leaving the local-header sizes unknown (Size == -1).
            Stream sink = dataDescriptor ? new WriteOnlyNonSeekableStream(backing) : (Stream)backing;
            using (SysZip.ZipArchive zip = new SysZip.ZipArchive(sink, SysZip.ZipArchiveMode.Create, leaveOpen: true))
            {
                SysZip.ZipArchiveEntry e = zip.CreateEntry(entryName, SysZip.CompressionLevel.Optimal);
                using Stream es = e.Open();
                writeEntry(es);
            }
            return backing.ToArray();
        }

        /// <summary>Write-only, non-seekable pass-through used to force ZIP data descriptors.</summary>
        private sealed class WriteOnlyNonSeekableStream : Stream
        {
            private readonly Stream _inner;
            public WriteOnlyNonSeekableStream(Stream inner) { _inner = inner; }
            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => throw new NotSupportedException();
            public override long Position
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }
            public override void Flush() => _inner.Flush();
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
        }
    }
}
