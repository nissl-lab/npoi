using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.Util
{
    public class FilterOutputStream : OutputStream
    {
        protected OutputStream output;

        public override bool CanRead => throw new NotImplementedException();

        public override bool CanSeek => throw new NotImplementedException();

        public override bool CanWrite => throw new NotImplementedException();

        public override long Length => throw new NotImplementedException();

        public override long Position { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public FilterOutputStream(OutputStream output)
        {
            this.output = output;
        }
        public override void Write(int b)
        {
            output.Write(b);
        }
        public override void Write(byte[] b)
        {
            Write(b, 0, b.Length);
        }
        public override void Write(byte[] b, int off, int len)
        {
            if ((off | len | (b.Length - (len + off)) | (off + len)) < 0)
                throw new IndexOutOfRangeException();

            for (int i = 0; i < len; i++)
            {
                Write(b[off + i]);
            }
        }
        public override void Flush()
        {
            output.Flush();
        }
        public override void Close()
        {
            Flush();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotImplementedException();
        }

        public override void SetLength(long value)
        {
            throw new NotImplementedException();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotImplementedException();
        }
    }
}
