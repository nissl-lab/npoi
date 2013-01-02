/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using System;
using System.IO;

namespace NPOI.POIFS.FileSystem
{
    internal class CloseIgnoringInputStream : Stream
    {

        private Stream _is;

        public CloseIgnoringInputStream(Stream stream)
        {
            _is = stream;
        }
        public int Read()
        {
            return (int)_is.ReadByte();
        }
        public override int Read(byte[] b, int off, int len)
        {
            return _is.Read(b, off, len);
        }
        public override void Close()
        {
            // do nothing
        }
        public override void Flush()
        {
        }
        public override long Seek(long offset, SeekOrigin origin)
        {
            return 0L;
        }

        public override void SetLength(long value)
        {
        }
        // Properties
        public override bool CanRead
        {
            get
            {
                return true;
            }
        }

        public override bool CanSeek
        {
            get
            {
                return true;
            }
        }

        public override bool CanWrite
        {
            get
            {
                return false;
            }
        }

        public override long Length
        {
            get
            {
                return (long)this._is.Length;
            }
        }

        public override long Position
        {
            get
            {
                return (long)this._is.Position;
            }
            set
            {
                this._is.Position = Convert.ToInt32(value);
            }
        }
        public override void Write(byte[] buffer, int offset, int count)
        {
        }
    }
}
