using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.Util
{
    public abstract class OutputStream : Stream
    {
        public abstract void Write(int b);
        public virtual void Write(byte[] b)
        {
            Write(b, 0, b.Length);
        }
        public override void Write(byte[] b, int off, int len)
        {
            if (b == null)
            {
                throw new NullReferenceException();
            }
            else if ((off < 0) || (off > b.Length) || (len < 0) ||
                         ((off + len) > b.Length) || ((off + len) < 0))
            {
                throw new IndexOutOfRangeException();
            }
            else if (len == 0)
            {
                return;
            }
            for (int i = 0; i < len; i++)
            {
                Write(b[off + i]);
            }
        }

        // ---- Stream abstract members / capabilities ----
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() { /* no-op by default */ }

        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();
    }
}
