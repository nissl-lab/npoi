using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.Util
{
    public class ByteArrayOutputStream : OutputStream
    {
        private List<byte> _buffer;
        private long _position;

        public ByteArrayOutputStream()
            : this(32)
        {
        }

        public ByteArrayOutputStream(int size) //: base(size)
        {
            _buffer = new List<byte>(size);
            _position = 0;
        }

        public override bool CanRead => false;
        public override bool CanSeek => true;
        public override bool CanWrite => true;

        public override long Length => _buffer.Count;

        public override long Position
        {
            get => _position;
            set
            {
                if (value < 0 || value > _buffer.Count)
                    throw new ArgumentOutOfRangeException(nameof(value));
                _position = value;
            }
        }

        public override void Flush()
        {
            // No-op for memory stream
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException("Read is not supported on ByteArrayOutputStream.");
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            long newPos;
            switch (origin)
            {
                case SeekOrigin.Begin:
                    newPos = offset;
                    break;
                case SeekOrigin.Current:
                    newPos = _position + offset;
                    break;
                case SeekOrigin.End:
                    newPos = _buffer.Count + offset;
                    break;
                default:
                    throw new ArgumentException("Invalid SeekOrigin", nameof(origin));
            }
            if (newPos < 0 || newPos > _buffer.Count)
                throw new IOException("Attempted to seek outside the buffer.");
            _position = newPos;
            return _position;
        }

        public override void SetLength(long value)
        {
            if (value < 0)
                throw new ArgumentOutOfRangeException(nameof(value));
            if (value < _buffer.Count)
            {
                _buffer.RemoveRange((int)value, _buffer.Count - (int)value);
            }
            else if (value > _buffer.Count)
            {
                _buffer.AddRange(new byte[value - _buffer.Count]);
            }
            if (_position > value)
                _position = value;
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            if (buffer == null)
                throw new ArgumentNullException(nameof(buffer));
            if (offset < 0 || count < 0 || offset + count > buffer.Length)
                throw new ArgumentOutOfRangeException();
            for (int i = 0; i < count; i++)
            {
                WriteByte(buffer[offset + i]);
            }
        }

        public override void Write(int b)
        {
            WriteByte((byte)b);
        }

        public override void WriteByte(byte value)
        {
            if (_position < _buffer.Count)
            {
                _buffer[(int)_position] = value;
            }
            else
            {
                _buffer.Add(value);
            }
            _position++;
        }

        public virtual void Write(byte[] b)
        {
            Write(b, 0, b.Length);
        }

        public void Reset()
        {
            _position = 0;
        }

        public byte[] ToByteArray()
        {
            return _buffer.ToArray();
        }

        public string ToString(string encoding)
        {
            return Encoding.GetEncoding(encoding).GetString(_buffer.ToArray());
        }
    }
}
