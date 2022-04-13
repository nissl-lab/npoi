using System;
using System.IO;

namespace NPOI.OpenXml4Net.Util
{
    public class VirtualStream : Stream
    {
        private Stream _stream;
        private long _begin;
        private long _end;

        public override bool CanRead { get { return _stream.CanRead; } }
        public override bool CanSeek { get { return _stream.CanSeek; } }
        public override bool CanWrite { get { return false; } }
        public override long Length
        {
            get
            {
                return _end - _begin;
            }
        }
        public override long Position
        {
            get
            {
                return _stream.Position - _begin;
            }
            set
            {
                _stream.Position = value + _begin;
            }
        }

        public VirtualStream(Stream stream, long begin, long end)
        {
            _stream = stream;
            _begin = begin;
            _end = end;
            Position = 0;
        }

        public override void Flush()
        {
            throw new NotImplementedException();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            var rst = _stream.Read(buffer, offset, count);
            return rst;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return _stream.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            throw new NotImplementedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotImplementedException();
        }

        private string ConvertBytes(byte[] bytes)
        {
            string strTemp = System.BitConverter.ToString(bytes);
            string[] strSplit = strTemp.Split('-');
            byte[] bytTemp2 = new byte[strSplit.Length];
            for (int i = 0; i < strSplit.Length; i++)
            {
                bytTemp2[i] = byte.Parse(strSplit[i], System.Globalization.NumberStyles.AllowHexSpecifier);
            }
            string strResult = System.Text.Encoding.Default.GetString(bytTemp2);
            return strResult;
        }
    }
}
