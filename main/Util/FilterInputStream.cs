using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.Util
{
    public class FilterInputStream : InputStream
    {
        protected volatile InputStream input;

        public override bool CanRead
        {
            get { throw new NotImplementedException(); }
        }

        public override bool CanSeek
        {
            get { throw new NotImplementedException(); }
        }

        public override bool CanWrite
        {
            get { throw new NotImplementedException(); }
        }

        public override long Length
        {
            get { throw new NotImplementedException(); }
        }

        public override long Position { get; set; }

        protected FilterInputStream(InputStream input)
        {
            this.input = input;
        }
        public override int Read()
        {
            return input.Read();
        }
        public override int Read(byte[] b)
        {
            return Read(b, 0, b.Length);
        }
        public override int Read(byte[] b, int off, int len)
        {
            return input.Read(b, off, len);
        }
        public override long Skip(long n)
        {
            return input.Skip(n);
        }
        public int available()
        {
            return input.Available();
        }
        public override void Close()
        {
            input.Close();
        }
        public override void Mark(int readlimit)
        {
            input.Mark(readlimit);
        }
        public override void Reset()
        {
            input.Reset();
        }
        public override bool MarkSupported()
        {
            return input.MarkSupported();
        }

        public override void Flush()
        {
            throw new NotImplementedException();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotImplementedException();
        }

        public override void SetLength(long value)
        {
            throw new NotImplementedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotImplementedException();
        }
    }
}
