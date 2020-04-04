using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.Util
{
    public class BoundedInputStream : ByteArrayInputStream
    {
        //the wrapped input stream
        private ByteArrayInputStream in1;

        //the max length to provide
        private long max;

        //the number of bytes already returned
        //private long pos = 0;

        //the marked position
        //private long mark = -1;

        // flag if close shoud be propagated
        //private bool propagateClose = true;
        /// <summary>
        /// Creates a new <c>BoundedInputStream</c> that wraps the given input
        /// stream and limits it to a certain size.
        /// </summary>
        /// <param name="in1">The wrapped input stream</param>
        /// <param name="size">The maximum number of bytes to return</param>
        public BoundedInputStream(ByteArrayInputStream in1, long size)
        {
            IsPropagateClose = true;
            // Some badly designed methods - eg the servlet API - overload length
            // such that "-1" means stream finished
            this.max = size;
            this.in1 = in1;
        }
        /// <summary>
        /// Creates a new <c>BoundedInputStream</c> that wraps the given input
        /// stream and is unlimited.
        /// </summary>
        /// <param name="in1">The wrapped input stream</param>
        public BoundedInputStream(ByteArrayInputStream in1)
            : this(in1, -1)
        {
            ;
        }
        /// <summary>
        /// Invokes the delegate's <code>read()</code> method if
        /// the current position is less than the limit.
        /// </summary>
        /// <returns>the byte read or -1 if the end of stream 
        /// or the limit has been reached.</returns>
        /// <exception cref="IOException">if an I/O error occurs</exception>
        public override int Read()
        {
            if (max >= 0 && pos == max)
            {
                return -1;
            }
            int result = in1.Read();
            pos++;
            return result;
        }
        public override int Read(byte[] b)
        {
            return this.Read(b, 0, b.Length);
        }
        public override int Read(byte[] b, int off, int len)
        {
            if (max >= 0 && pos >= max)
            {
                return -1;
            }
            long maxRead = max >= 0 ? Math.Min(len, max - pos) : len;
            int bytesRead = in1.Read(b, off, (int)maxRead);

            if (bytesRead == -1)
            {
                return -1;
            }

            pos += bytesRead;
            return bytesRead;
        }
        
        public override void Close()
        {
            if (IsPropagateClose)
            {
                in1.Close();
            }
        }
        public override void Reset()
        {
            in1.Reset();
            pos = mark;
        }
        public override void Mark(int readlimit)
        {
            in1.Mark(readlimit);
            mark = pos;
        }
        public override bool MarkSupported()
        {
            return in1.MarkSupported();
        }
        public bool IsPropagateClose { get; set; }
    }
}
