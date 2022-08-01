using System;
using System.IO;

namespace NPOI.Util
{
    public class FileInputStream : InputStream
    {
        readonly Stream inner;

        public FileInputStream(Stream fs)
        {
            this.inner = fs;
        }

        public override bool CanRead
        {
            get { return inner.CanRead; }
        }

        public override bool CanSeek
        {
            get { return false; }
        }

        public override bool CanWrite
        {
            get { return false; }
        }

        public override long Length
        {
            get { return inner.Length; }
        }

        public override long Position
        {
            get { return inner.Position;}
            set { inner.Position = value;}
        }

        public override void Flush()
        {
            throw new NotImplementedException();
        }

        public override int Read()
        {
            return inner.ReadByte();
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

        public override void Close()
        {
            if (inner != null)
                inner.Close();
        }
    }
}
