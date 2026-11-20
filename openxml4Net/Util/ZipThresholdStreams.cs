using System;
using System.IO;

namespace NPOI.OpenXml4Net.Util
{
    /// <summary>
    /// A read-only pass-through stream that counts the number of raw (compressed)
    /// bytes consumed from the underlying stream.
    /// <para>
    /// This is used together with <see cref="ZipSecureFile.CheckThreshold"/> to detect
    /// decompression bombs by comparing the amount of compressed data actually read
    /// against the amount of decompressed data produced. It works even when a ZIP
    /// entry's local header omits the sizes (streamed / data-descriptor entries, where
    /// <c>ZipEntry.Size == -1</c>), because the ratio is computed from bytes observed
    /// on the wire rather than from a possibly-absent header field.
    /// </para>
    /// </summary>
    public sealed class CountingStream : Stream
    {
        private readonly Stream _inner;
        private long _bytesRead;

        public CountingStream(Stream inner)
        {
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        }

        /// <summary>Total number of bytes read from the underlying stream so far.</summary>
        public long BytesRead
        {
            get { return _bytesRead; }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            int n = _inner.Read(buffer, offset, count);
            if (n > 0)
            {
                _bytesRead += n;
            }
            return n;
        }

        public override int ReadByte()
        {
            int b = _inner.ReadByte();
            if (b >= 0)
            {
                _bytesRead++;
            }
            return b;
        }

        public override bool CanRead { get { return _inner.CanRead; } }
        public override bool CanSeek { get { return _inner.CanSeek; } }
        public override bool CanWrite { get { return false; } }
        public override long Length { get { return _inner.Length; } }

        public override long Position
        {
            get { return _inner.Position; }
            set { _inner.Position = value; }
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) { return _inner.Seek(offset, origin); }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _inner.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// A read-only, forward-only wrapper around the <em>decompressed</em> stream of a
    /// single ZIP entry. As bytes are read it accumulates the decompressed size and
    /// validates it against the zip-bomb thresholds via
    /// <see cref="ZipSecureFile.CheckThreshold"/>, using the entry's known compressed
    /// size for the inflate-ratio check.
    /// <para>
    /// Used on the seekable (<see cref="ZipFileZipEntrySource"/>) read path, where the
    /// central-directory compressed size is reliable.
    /// </para>
    /// </summary>
    public sealed class ZipEntrySizeGuardStream : Stream
    {
        private readonly Stream _inner;
        private readonly long _compressedSize;
        private long _decompressed;

        public ZipEntrySizeGuardStream(Stream inner, long compressedSize)
        {
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            _compressedSize = compressedSize;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            int n = _inner.Read(buffer, offset, count);
            if (n > 0)
            {
                _decompressed += n;
                ZipSecureFile.CheckThreshold(_decompressed, _compressedSize);
            }
            return n;
        }

        public override int ReadByte()
        {
            int b = _inner.ReadByte();
            if (b >= 0)
            {
                _decompressed++;
                ZipSecureFile.CheckThreshold(_decompressed, _compressedSize);
            }
            return b;
        }

        public override bool CanRead { get { return _inner.CanRead; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return false; } }

        // The underlying decompressed (inflater) stream is forward-only; mirror that so
        // callers cannot seek past the guard. Position reports bytes produced so far.
        public override long Length { get { throw new NotSupportedException(); } }

        public override long Position
        {
            get { return _decompressed; }
            set { throw new NotSupportedException(); }
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _inner.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
