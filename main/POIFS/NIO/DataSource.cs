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

using System.IO;
using NPOI.Util;
namespace NPOI.POIFS.NIO
{
    /// <summary>
    /// Common definition of how we read and write bytes
    /// </summary>
    public abstract class DataSource
    {
        public abstract ByteBuffer Read(int length, long position);

        public abstract void Write(ByteBuffer src, long position);
        public abstract long Size { get; }
        /// <summary>
        /// Close the underlying stream
        /// </summary>
        public abstract void Close();
        /// <summary>
        /// Copies the contents to the specified Stream
        /// </summary>
        /// <param name="stream"></param>
        public abstract void CopyTo(Stream stream);
    }
}