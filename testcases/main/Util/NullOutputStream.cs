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

using NPOI.Util;
using System;
using System.IO;

namespace TestCases.Util
{
    public class NullOutputStream : OutputStream
    {
        public NullOutputStream()
        {
        }
        public override bool CanRead
        {
            get { return false;}
        }
        public override bool CanSeek
        {
            get { return true;}
        }
        public override bool CanWrite
        {
            get { return true;}
        }
        private long _length;
        public override long Length
        {
            get { return _length; }
        }

        public override long Position { get; set; }

        public override void Flush()
        {
            
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return 0;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return 0;
        }

        public override void SetLength(long value)
        {
            _length = value;
        }

        public override void Write(int b)
        {
        }
    }
}
