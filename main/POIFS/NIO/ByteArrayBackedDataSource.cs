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
using System;
using NPOI.Util;

namespace NPOI.POIFS.NIO
{
    /// <summary>
    /// A POIFS <see cref="T:NPOI.POIFS.NIO.DataSource"/> backed by a byte array.
    /// </summary>
    public class ByteArrayBackedDataSource : DataSource
    {
        private byte[] buffer;
        private long size;
        
        public ByteArrayBackedDataSource(byte[] data, int size)
        {
            this.buffer = data;
            this.size = size;
        }
        public ByteArrayBackedDataSource(byte[] data)
            : this(data, data.Length)
        {
        }
        public override ByteBuffer Read(int length, long position)
        {
            if (position >= size)
            {
                throw new IndexOutOfRangeException(
                      "Unable to read " + length + " bytes from " +
                      position + " in stream of length " + size
                );
            }

            int toRead = (int)Math.Min(length, size - position);
            return ByteBuffer.CreateBuffer(buffer, (int)position, toRead);

        }


        public override void Write(ByteBuffer src, long position)
        {
            long endPosition = position + src.Length;

            if (endPosition > buffer.Length)
            {
                Extend(endPosition);
            }

            // Now copy
            src.Read(buffer, (int)position, src.Length);

            // Update size if needed
            if (endPosition > size)
            {
                size = endPosition;
            }
        }

        private void Extend(long length)
        {
            // Consider extending by a bit more than requested
            long difference = length - buffer.Length;
            if (difference < buffer.Length * 0.25)
            {
                difference = (long)(buffer.Length * 0.25);
            }
            if (difference < 4096)
            {
                difference = 4096;
            }
            byte[] nb = new byte[(int)(difference + buffer.Length)];
            Array.Copy(buffer, 0, nb, 0, (int)size);
            buffer = nb;
        }

        public override void CopyTo(Stream stream)
        {
            stream.Write(buffer, 0, (int)size);
        }

        public override long Size
        {
            get
            {
                return size;
            }
        }

        public override void Close()
        {
            buffer = null;
            size = -1;
        }
    }
}