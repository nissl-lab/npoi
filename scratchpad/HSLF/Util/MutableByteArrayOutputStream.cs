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

namespace NPOI.HSLF.Util
{
    using System;
    using System.IO;


    /**
     * This class doesn't work yet, but is here to show the idea of a
     *  MemoryStream where you can track how many bytes you've
     *  already written, and go back and write over a previous part of the stream
     *
     * @author Nick Burch
     */

    public class MutableMemoryStream : MemoryStream
    {
        /** Return how many bytes we've stuffed in so far */
        public int GetBytesWritten() { return -1; }

        /** Write some bytes to the array */
        public virtual void Write(byte[] b) { }
        public virtual void Write(int b) { }

        /** Write some bytes to an earlier bit of the array */
        public void OverWrite(byte[] b, int startPos) { }
    }
}
