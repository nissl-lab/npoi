using NPOI.POIFS.Crypt;
using System;
using System.IO;

namespace NPOI.POIFS.FileSystem
{
    /// <summary>
    ///     stream for NPOI IWorkbook with password
    /// </summary>
    public class PasswordXlsxOutputStream(string outputPath, string password) : Stream
    {
        private readonly MemoryStream _buffer = new();

        public override bool CanRead => false;
        public override bool CanSeek => true;
        public override bool CanWrite => true;
        public override long Length => _buffer.Length;

        public override long Position
        {
            get => _buffer.Position;
            set => _buffer.Position = value;
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            _buffer.Write(buffer, offset, count);
        }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return 0;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return _buffer.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            _buffer.SetLength(value);
        }

        public override void Close()
        {
            base.Close();
            var raw = _buffer.ToArray();
            XlsxEncryptor.FromBytesToFile(raw, outputPath, password);
        }

    }
}