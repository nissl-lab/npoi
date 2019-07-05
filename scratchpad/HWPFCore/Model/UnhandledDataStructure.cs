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

namespace NPOI.HWPF.Model
{
    using System;
    public class UnhandledDataStructure
    {
        byte[] _buf;

        public UnhandledDataStructure(byte[] buf, int offset, int length)
        {
            //    Console.WriteLine("Yes, using my code");
            _buf = new byte[length];
            if (offset + length > buf.Length)
            {
                throw new IndexOutOfRangeException("buffer length is " + buf.Length +
                                                    "but code is trying to read " + length + " from offset " + offset);
            }
            Array.Copy(buf, offset, _buf, 0, length);
        }

        internal byte[] GetBuf()
        {
            return _buf;
        }
    }
}


