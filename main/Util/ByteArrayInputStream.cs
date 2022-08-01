using System;
using System.IO;

namespace NPOI.Util
{
    public class ByteArrayInputStream : InputStream
    {
        protected byte[] buf;
        protected int pos;
        protected int mark = 0;
        protected int count;

        private readonly object _lockObject = new object();

        public ByteArrayInputStream()
        {
        }

        public ByteArrayInputStream(byte[] buf)
        {
            this.buf = buf;
            this.pos = 0;
            this.count = buf.Length;
        }

        public ByteArrayInputStream(byte[] buf, int offset, int length)
        {
            this.buf = buf;
            this.pos = offset;
            this.count = Math.Min(offset + length, buf.Length);
            this.mark = offset;
        }

        public override int Read()
        {
            lock (_lockObject)
            {
                return (pos < count) ? (buf[pos++] & 0xff) : -1;
            }
        }

        public override int Read(byte[] b, int off, int len)
        {
            lock (_lockObject)
            {
                if (b == null)
                {
                    throw new NullReferenceException();
                }
                else if (off < 0 || len < 0 || len > b.Length - off)
                {
                    throw new IndexOutOfRangeException();
                }

                if (pos >= count)
                {
                    return 0;
                }

                int avail = count - pos;
                if (len > avail)
                {
                    len = avail;
                }

                if (len <= 0)
                {
                    return 0;
                }

                Array.Copy(buf, pos, b, off, len);
                pos += len;
                return len;
            }
        }

        public override int Available()
        {
            return count - pos;
        }

        public override bool MarkSupported()
        {
            return true;
        }

        public override void Mark(int readlimit)
        {
            mark = pos;
        }

        public override void Reset()
        {
            pos = mark;
        }

        public override void Close() { }

        public override bool CanRead
        {
            get { return true; }
        }

        public override bool CanWrite
        {
            get { return false; }
        }

        public override bool CanSeek
        {
            get { return true; }
        }

        public override void Flush()
        {
            throw new NotImplementedException();
        }

        public override long Length
        {
            get { return this.count; }
        }

        public override long Position
        {
            get {  return this.pos; }
            set { this.pos = (int)value; }
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            if (!this.CanSeek)
                throw new NotSupportedException();

            switch (origin)
            {
                case SeekOrigin.Begin:
                    if (0L > offset)
                    {
                        throw new ArgumentOutOfRangeException(nameof(offset), "offset must be positive");
                    }
                    this.Position = offset < this.Length ? offset : this.Length;
                    break;

                case SeekOrigin.Current:
                    this.Position = (this.Position + offset) < this.Length ? (this.Position + offset) : this.Length;
                    break;

                case SeekOrigin.End:
                    this.Position = this.Length;
                    break;

                default:
                    throw new ArgumentException("incorrect SeekOrigin", nameof(origin));
            }
            return Position;
        }

        public override void SetLength(long value)
        {
            throw new NotImplementedException();
        }
    }
}
