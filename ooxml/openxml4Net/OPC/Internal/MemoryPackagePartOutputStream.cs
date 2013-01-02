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
            _buff = new MemoryStream();
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
            if (_part.data != null)
            {
                //byte[] newArray = new byte[_part.data.Length + _buff.Length];
                //// copy the previous contents of part.data in newArray
                //Array.Copy(_part.data, 0, newArray, 0, _part.data.Length);

                //// append the newly added data
                //byte[] buffArr = _buff.ToArray();
                //Array.Copy(buffArr, 0, newArray, _part.data.Length,
                //        buffArr.Length);

                byte[] newArray = new byte[_buff.Length];
                byte[] buffArr = _buff.ToArray();
                Array.Copy(buffArr, 0, newArray, 0,
                        buffArr.Length);

                // save the result as new data
                _part.data = newArray;
            }
            else
            {
                // was empty, just fill it
                _part.data = _buff.ToArray();
            }

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
