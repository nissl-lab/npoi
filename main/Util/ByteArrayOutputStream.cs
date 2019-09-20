using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.Util
{
    public class ByteArrayOutputStream : Stream
    {
        protected byte[] buf;
        protected int count;
        public ByteArrayOutputStream()
            : this(32)
        {

        }
        public ByteArrayOutputStream(int size)
        {
            if (size < 0)
            {
                throw new ArgumentOutOfRangeException("Negative initial size: "
                                                   + size);
            }
            buf = new byte[size];
        }

        private void EnsureCapacity(int minCapacity)
        {
            // overflow-conscious code
            if (minCapacity - buf.Length > 0)
                Grow(minCapacity);
        }
        private static int MAX_ARRAY_SIZE = Int32.MaxValue - 8;

        public override bool CanRead => throw new NotImplementedException();

        public override bool CanSeek => throw new NotImplementedException();

        public override bool CanWrite => throw new NotImplementedException();

        public override long Length => throw new NotImplementedException();

        public override long Position { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        private void Grow(int minCapacity)
        {
            // overflow-conscious code
            int oldCapacity = buf.Length;
            int newCapacity = oldCapacity << 1;
            if (newCapacity - minCapacity < 0)
                newCapacity = minCapacity;
            if (newCapacity - MAX_ARRAY_SIZE > 0)
                newCapacity = HugeCapacity(minCapacity);
            buf = Arrays.CopyOf(buf, newCapacity);
        }

        private static int HugeCapacity(int minCapacity)
        {
            if (minCapacity < 0) // overflow
                throw new OutOfMemoryException();
            return (minCapacity > MAX_ARRAY_SIZE) ?
                Int32.MaxValue :
                MAX_ARRAY_SIZE;
        }
        public void Write(int b)
        {
            EnsureCapacity(count + 1);
            buf[count] = (byte)b;
            count += 1;
        }
        public void Write(byte[] b)
        {
            Write(b, 0, b.Length);
        }
        public override void Write(byte[] b, int off, int len)
        {
            if ((off < 0) || (off > b.Length) || (len < 0) ||
                ((off + len) - b.Length > 0))
            {
                throw new IndexOutOfRangeException();
            }
            EnsureCapacity(count + len);
            Array.Copy(b, off, buf, count, len);
            count += len;
        }
        public void WriteTo(Stream out1)
        {
            out1.Write(buf, 0, count);
        }
        public void Reset()
        {
            count = 0;
        }
        public byte[] ToByteArray()
        {
            return Arrays.CopyOf(buf, count);
        }
        public int Size()
        {
            return count;
        }
        public override string ToString()
        {
            char[] cs = new char[count];
            Array.Copy(buf, 0, cs, 0, count);
            return new String(cs, 0, count);
        }
        public String ToString(String charsetName)
        {
            return Encoding.GetEncoding(charsetName).GetString(buf);
            //return new String(buf, 0, count, charsetName);
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

        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotImplementedException();
        }
    }
}
