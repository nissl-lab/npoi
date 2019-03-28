using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace NPOI.OpenXml4Net.OPC.Internal
{
    public class MemoryPackagePartOutputStream : Stream
    {
        private MemoryPackagePart _part;

        private MemoryStream _buff;

        public MemoryPackagePartOutputStream(MemoryPackagePart part)
        {
            this._part = part;
            //if (this._part.data == null)
            {
                this._part.data = new MemoryStream();
            }
            _buff = this._part.data;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotImplementedException();
        }

        public override bool CanRead
        {
            get { return false; }
        }
        public override bool CanWrite
        {
            get { return true; }
        }
        public override bool CanSeek
        {
            get { return false; }
        }

        public override long Length
        {
            get { return _buff.Length; }
        }
        
        public void Write(int b)
        {
            _buff.WriteByte((byte)b);
        }
        public override void SetLength(long value)
        {
            _buff.SetLength(value);
        }
        public override long Seek(long offset, SeekOrigin origin)
        {
            return _buff.Seek(offset, origin);
        }
        public override long Position
        {
            get
            {
                return _buff.Position;
            }
            set
            {
                _buff.Position = value;
            }
        }

        /**
         * Close this stream and flush the content.
         * @see #flush()
         */
        public override void Close()
        {
            this.Flush();
        }

        /**
         * Flush this output stream. This method is called by the close() method.
         * Warning : don't call this method for output consistency.
         * @see #close()
         */
        public override void Flush()
        {
            _buff.Flush();

            /*
             * Clear this streams buffer, in case flush() is called a second time
             * Fix bug 1921637 - provided by Rainer Schwarze
             */
            _buff.Position = 0;
        }


        public override void Write(byte[] b, int off, int len)
        {
            _buff.Write(b, off, len);
        }


        public void Write(byte[] b)
        {
            _buff.Write(b, (int)_buff.Position, b.Length);
        }
    }
}
