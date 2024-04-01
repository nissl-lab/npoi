/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

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

        /// <summary>
        /// Close this stream and flush the content.
        /// <see cref="flush()" />
        /// </summary>
        public override void Close()
        {
            this.Flush();
        }

        /// <summary>
        /// Flush this output stream. This method is called by the close() method.
        /// Warning : don't call this method for output consistency.
        /// <see cref="close()" />
        /// </summary>
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
