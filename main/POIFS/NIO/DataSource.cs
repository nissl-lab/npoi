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
using System.Threading;
using System.Threading.Tasks;
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

        /// <summary>
        /// Copies the contents to the specified Stream asynchronously
        /// </summary>
        /// <param name="stream">The stream to copy to</param>
        /// <param name="cancellationToken">Cancellation token to observe during the async operation</param>
        /// <returns>A task that represents the asynchronous copy operation</returns>
        public virtual async Task CopyToAsync(Stream stream, CancellationToken cancellationToken = default)
        {
            // Default implementation calls synchronous CopyTo
            // Most subclasses override this with true async implementation
            // This is kept for backward compatibility only
            await Task.Yield(); // Allow other tasks to run
            cancellationToken.ThrowIfCancellationRequested();
            CopyTo(stream);
        }
    }
}