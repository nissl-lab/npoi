using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace NPOI.Util
{
    public class ByteArrayInputStream : MemoryStream
    {
        public ByteArrayInputStream()
        {
        }
        protected byte[] buf;
        protected int pos;
        protected int mark = 0;
        protected int count;
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
        
        public virtual int Read()
        {
            lock (this)
            {
                return (pos < count) ? (buf[pos++] & 0xff) : -1;
            }
        }
        public override int Read(byte[] b, int off, int len)
        {
            lock (this)
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
                    return -1;
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
        public virtual int Available()
        {
            return count - pos;
        }
        public virtual bool MarkSupported()
        {
            return true;
        }
        public virtual void Mark(int readAheadLimit)
        {
            mark = pos;
        }
        public virtual void Reset()
        {
            pos = mark;
        }
        public override void Close()
        {
        }
    }
}
