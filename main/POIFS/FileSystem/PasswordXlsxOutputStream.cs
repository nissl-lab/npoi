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
 * Contributors: Seijiro Ikehata (modeverv at gmail.com)
 *
 * ==============================================================*/

namespace NPOI.POIFS.FileSystem
{
    using NPOI.POIFS.Crypt;
    using System;
    using System.IO;

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