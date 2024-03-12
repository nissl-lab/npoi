using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.Util
{
    public class ByteArrayOutputStream : MemoryStream
    {
        public ByteArrayOutputStream()
            : this(32)
        {

        }
        public ByteArrayOutputStream(int size):base(size)
        {

        }
        public virtual void Write(int b)
        {
            WriteByte((byte)b);
        }
        public virtual void Write(byte[] b)
        {
            Write(b, 0, b.Length);
        }
        public void Reset()
        {
            this.Position = 0;
        }
        public byte[] ToByteArray()
        {
            return this.ToArray();
        }
        public string ToString(string encoding)
        {
            return Encoding.GetEncoding(encoding).GetString(this.ToArray());
        }
    }
}
