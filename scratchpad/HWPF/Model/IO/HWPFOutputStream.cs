
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HWPF.Model.IO
{
    using System;
    using System.IO;
    using System.Collections;

    public class HWPFStream : MemoryStream
    {

        int _offset;

        public HWPFStream()
            : base()
        {

        }

        public int Offset
        {
            get
            {
                return _offset;
            }
        }

        public void Reset()
        {
            _offset = 0;
        }

        public override void Write(byte[] buf, int off, int len)
        {
            base.Write(buf, off, len);
            _offset += len;
        }

        public void Write(byte[] buf)
        {
            base.Write(buf, 0, buf.Length);
            _offset += buf.Length;
        }

        public void Write(int b)
        {
            base.WriteByte((byte)b);
            _offset++;
        }
    }
}