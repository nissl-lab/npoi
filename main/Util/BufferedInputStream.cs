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

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Util
{
    public class BufferedInputStream : FilterInputStream
    {
        private static int DEFAULT_BUFFER_SIZE = 8192;
        private static int MAX_BUFFER_SIZE = int.MaxValue - 8;
        protected volatile byte[] buf;
        protected int count;
        protected int pos;
        protected int markpos = -1;
        protected int marklimit;
        public BufferedInputStream(InputStream input)
            : this(input, DEFAULT_BUFFER_SIZE)
        {
        }
        public BufferedInputStream(InputStream input, int size)
            : base(input)
        {
            if (size <= 0)
            {
                throw new ArgumentException("Buffer size <= 0");
            }
            buf = new byte[size];
        }

        public override int Read(byte[] b, int off, int len)
        {
            //getBufIfOpen(); // Check for closed stream
            if ((off | len | (off + len) | (b.Length - (off + len))) < 0) {
                throw new IndexOutOfRangeException();
            } else if (len == 0) {
                return 0;
            }

            int n = 0;
            for (;;) {
                int nread = Read1(b, off + n, len - n);
                if (nread <= 0)
                    return (n == 0) ? nread : n;
                n += nread;
                if (n >= len)
                    return n;
                // if not closed but no bytes available, return
                InputStream input = this.input;
                if (input != null && input.Available() <= 0)
                    return n;
            }
        }

        private int Read1(byte[] b, int off, int len) {
            int avail = count - pos;
            if (avail <= 0) {
                /* If the requested length is at least as large as the buffer, and
                   if there is no mark/reset activity, do not bother to copy the
                   bytes into the local buffer.  In this way buffered streams will
                   cascade harmlessly. */
                if (len >= buf.Length && markpos < 0) {
                    return input.Read(b, off, len);
                }
                Fill();
                avail = count - pos;
                if (avail <= 0) return -1;
            }
            int cnt = (avail < len) ? avail : len;
            Array.Copy(buf, pos, b, off, cnt);
            pos += cnt;
            return cnt;
        }
        private void Fill()
        {
            byte[] buffer = buf;
            if (markpos < 0)
                pos = 0;            /* no mark: throw away the buffer */
            else if (pos >= buffer.Length)  /* no room left in buffer */
                if (markpos > 0) {  /* can throw away early part of the buffer */
                    int sz = pos - markpos;
                    Array.Copy(buffer, markpos, buffer, 0, sz);
                    pos = sz;
                    markpos = 0;
                } else if (buffer.Length >= marklimit) {
                    markpos = -1;   /* buffer got too big, invalidate mark */
                    pos = 0;        /* drop buffer contents */
                } else if (buffer.Length >= MAX_BUFFER_SIZE) {
                    throw new OutOfMemoryException("Required array size too large");
                } else {            /* grow buffer */
                    int nsz = (pos <= MAX_BUFFER_SIZE - pos) ?
                            pos * 2 : MAX_BUFFER_SIZE;
                    if (nsz > marklimit)
                        nsz = marklimit;
                    byte[] nbuf = new byte[nsz];
                    Array.Copy(buffer, 0, nbuf, 0, pos);
                    //if (!bufUpdater.compareAndSet(this, buffer, nbuf)) {
                    //    // Can't replace buf if there was an async close.
                    //    // Note: This would need to be changed if fill()
                    //    // is ever made accessible to multiple threads.
                    //    // But for now, the only way CAS can fail is via close.
                    //    // assert buf == null;
                    //    throw new IOException("Stream closed");
                    //}
                    this.buf = nbuf;
                    buffer = nbuf;
                }
            count = pos;
            int n = input.Read(buffer, pos, buffer.Length - pos);
            if (n > 0)
                count = n + pos;
        }

        public override int Available()
        {
            int n = count - pos;
            int avail = input.Available();
            return n > (int.MaxValue - avail) ? int.MaxValue : n + avail;
        }

        public override void Mark(int readlimit) {
            marklimit = readlimit;
            markpos = pos;
        }

        public override bool MarkSupported() {
            return true;
        }

        public override void Reset()
        {
            //getBufIfOpen(); // Cause exception if closed
            if (markpos < 0)
                throw new IOException("Resetting to invalid mark");
            pos = markpos;
        }
    }
}
