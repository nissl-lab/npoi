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
using System.Text;

namespace TestCases.POIFS.FileSystem
{
    /**
     * An input stream which will return a maximum of
     *  a given number of bytes to Read, and often claims
     *  not to have any data
     */
    internal class SlowInputStream : Stream
    {
        private Random r = new Random(0);
        private byte[] data;
        private int chunkSize;
        private long pos = 0;

        public SlowInputStream(MemoryStream stream)
        {
            this.data = stream.ToArray();
            this.chunkSize = 512;    //default value
            this.pos = stream.Position;
        }

        public SlowInputStream(byte[] data, int chunkSize)
        {
            this.chunkSize = chunkSize;
            this.data = data;
        }

        /**
         * 75% of the time, claim there's no data available
         */
        private bool ClaimNoData()
        {
            double tmp = r.NextDouble();
            if (tmp < 0.75f)   //change 0.25f to 0.40f
            {
                return false;
            }
            return true;
        }

        public int Read()
        {
            if (pos >= data.Length)
            {
                return -1;
            }
            int ret = data[pos];
            pos++;

            if (ret < 0)
                ret += 256;

            return ret;
        }

        /**
         * Reads the requested number of bytes, or the chunk
         *  size, whichever Is lower.
         * Quite often will simply claim to have no data
         */
        public override int Read(byte[] b, int off, int len)
        {
            // Keep the Length within the chunk size
            if (len > chunkSize)
            {
                len = chunkSize;
            }
            // Don't Read off the end of the data
            if (pos + len > data.Length)
            {
                len = data.Length - (int)pos;

                // Spot when we're out of data
                if (len == 0)
                {
                    return -1;
                }
            }

            // Copy, and return what we Read
            Array.Copy(data, pos, b, off, len);
            pos += len;
            return len;
        }

        public int Read(byte[] b)
        {
            return Read(b, 0, b.Length);
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
                return data.Length;
            }
        }

        public override long Position
        {
            get
            {
                return pos;
            }
            set
            {
                pos = value;
            }
        }


        public override void Write(byte[] buffer, int offset, int count)
        {
        }
    }
}
