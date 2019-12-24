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


namespace NPOI.Util
{
    using System;
    using System.IO;

	public class PushbackInputStream : FilterInputStream
    {
        protected byte[] buf;
        private int bufint = -1;
        protected int pos;
        public PushbackInputStream(InputStream input): this(input, 1)
		{
		}
        public PushbackInputStream(InputStream input, int size)
            : base(input)
        {
            if (size <= 0)
            {
                throw new ArgumentException("size <= 0");
            }
            this.buf = new byte[size];
            this.pos = size;
        }
        protected override void Dispose(bool disposing)
        {
            this.input = null;

            base.Dispose(disposing);
        }

        /// <summary>
        /// Reads a byte from the stream and advances the position within the stream by one byte, or returns -1 if at the end of the stream.
        /// </summary>
        /// <returns>
        /// The unsigned byte cast to an Int32, or -1 if at the end of the stream.
        /// </returns>
        /// <exception cref="T:System.NotSupportedException">
        /// The stream does not support reading.
        /// </exception>
        /// <exception cref="T:System.ObjectDisposedException">
        /// Methods were called after the stream was closed.
        /// </exception>
		public override int ReadByte()
		{
			if (bufint != -1)
			{
				int tmp = bufint;
				bufint = -1;
				return tmp;
			}

			return input.ReadByte();
		}

        public override int Read()
        {
            ensureOpen();
            if (pos < buf.Length)
            {
                return buf[pos++] & 0xff;
            }
            return base.Read();
        }
        public override int Read(byte[] b, int off, int len)
        {
            ensureOpen();
            if (b == null)
            {
                throw new ArgumentNullException();
            }
            else if (off < 0 || len < 0 || len > b.Length - off)
            {
                throw new IndexOutOfRangeException();
            }
            else if (len == 0)
            {
                return 0;
            }

            int avail = buf.Length - pos;
            if (avail > 0)
            {
                if (len < avail)
                {
                    avail = len;
                }
                Array.Copy(buf, pos, b, off, avail);
                pos += avail;
                off += avail;
                len -= avail;
            }
            if (len > 0)
            {
                len = base.Read(b, off, len);
                if (len == -1)
                {
                    return avail == 0 ? -1 : avail;
                }
                return avail + len;
            }
            return avail;
        }
        /// <summary>
        /// When overridden in a derived class, reads a sequence of bytes from the current stream and advances the position within the stream by the number of bytes read.
        /// </summary>
        /// <param name="buffer">An array of bytes. When this method returns, the buffer contains the specified byte array with the values between <paramref name="offset"/> and (<paramref name="offset"/> + <paramref name="count"/> - 1) replaced by the bytes read from the current source.</param>
        /// <param name="offset">The zero-based byte offset in <paramref name="buffer"/> at which to begin storing the data read from the current stream.</param>
        /// <param name="count">The maximum number of bytes to be read from the current stream.</param>
        /// <returns>
        /// The total number of bytes read into the buffer. This can be less than the number of bytes requested if that many bytes are not currently available, or zero (0) if the end of the stream has been reached.
        /// </returns>
        /// <exception cref="T:System.ArgumentException">
        /// The sum of <paramref name="offset"/> and <paramref name="count"/> is larger than the buffer length.
        /// </exception>
        /// <exception cref="T:System.ArgumentNullException">
        /// 	<paramref name="buffer"/> is null.
        /// </exception>
        /// <exception cref="T:System.ArgumentOutOfRangeException">
        /// 	<paramref name="offset"/> or <paramref name="count"/> is negative.
        /// </exception>
        /// <exception cref="T:System.IO.IOException">
        /// An I/O error occurs.
        /// </exception>
        /// <exception cref="T:System.NotSupportedException">
        /// The stream does not support reading.
        /// </exception>
        /// <exception cref="T:System.ObjectDisposedException">
        /// Methods were called after the stream was closed.
        /// </exception>
  //      public override int Read(byte[] buffer, int offset, int count)
		//{
		//	if (bufint != -1 && count > 0)
		//	{
		//		// TODO Can this case be made more efficient?
		//		buffer[offset] = (byte) bufint;
		//		bufint = -1;
		//		return 1;
		//	}

		//	return input.Read(buffer, offset, count);
		//}

        /// <summary>
        /// Unreads the specified b.
        /// </summary>
        /// <param name="b">The b.</param>
		public virtual void Unread(int b)
		{
            ensureOpen();
            if (pos == 0)
            {
                throw new IOException("Push back buffer is full");
            }
            buf[--pos] = (byte)b;
        }
        public void Unread(byte[] b)
        {
            Unread(b, 0, b.Length);
        }

        public override int Available()
        {
            ensureOpen();
            int n = buf.Length - pos;
            int avail = base.Available();
            return n > (Int32.MaxValue - avail)
                        ? Int32.MaxValue
                        : n + avail;
        }
        /// <summary>
        /// When overridden in a derived class, gets a value indicating whether the current stream supports reading.
        /// </summary>
        /// <value></value>
        /// <returns>true if the stream supports reading; otherwise, false.
        /// </returns>
        public override bool CanRead
        {
            get { return input.CanRead; }
        }
        private void ensureOpen()
        {
            if (input == null)
                throw new IOException("Stream closed");
        }
        /// <summary>
        /// Pushes back a portion of an array of bytes by copying it to the front
        /// of the pushback buffer.
        /// </summary>
        /// <param name="b">the byte array to push back.</param>
        /// <param name="off">the start offset of the data.</param>
        /// <param name="len">the number of bytes to push back.</param>
        public void Unread(byte[] b, int off, int len)
        {
            ensureOpen();
            if (len > pos)
            {
                throw new IOException("Push back buffer is full");
            }
            pos -= len;
            //Position -= len;
            Array.Copy(b, off, buf, pos, len);
        }
        public override long Skip(long n)
        {
            ensureOpen();
            if (n <= 0)
            {
                return 0;
            }

            long pskip = buf.Length - pos;
            if (pskip > 0)
            {
                if (n < pskip)
                {
                    pskip = n;
                }
                pos += (int)pskip;
                n -= pskip;
            }
            if (n > 0)
            {
                pskip += base.Skip(n);
            }
            return pskip;
        }
        public override bool MarkSupported()
        {
            return false;
        }
        /// <summary>
        /// When overridden in a derived class, gets a value indicating whether the current stream supports seeking.
        /// </summary>
        /// <value></value>
        /// <returns>true if the stream supports seeking; otherwise, false.
        /// </returns>
        public override bool CanSeek
        {
            get { return input.CanSeek; }
        }
        /// <summary>
        /// When overridden in a derived class, gets a value indicating whether the current stream supports writing.
        /// </summary>
        /// <value></value>
        /// <returns>true if the stream supports writing; otherwise, false.
        /// </returns>
        public override bool CanWrite
        {
            get { return input.CanWrite; }
        }
        /// <summary>
        /// When overridden in a derived class, gets the length in bytes of the stream.
        /// </summary>
        /// <value></value>
        /// <returns>
        /// A long value representing the length of the stream in bytes.
        /// </returns>
        /// <exception cref="T:System.NotSupportedException">
        /// A class derived from Stream does not support seeking.
        /// </exception>
        /// <exception cref="T:System.ObjectDisposedException">
        /// Methods were called after the stream was closed.
        /// </exception>
        public override long Length
        {
            get { return input.Length; }
        }
        /// <summary>
        /// When overridden in a derived class, gets or sets the position within the current stream.
        /// </summary>
        /// <value></value>
        /// <returns>
        /// The current position within the stream.
        /// </returns>
        /// <exception cref="T:System.IO.IOException">
        /// An I/O error occurs.
        /// </exception>
        /// <exception cref="T:System.NotSupportedException">
        /// The stream does not support seeking.
        /// </exception>
        /// <exception cref="T:System.ObjectDisposedException">
        /// Methods were called after the stream was closed.
        /// </exception>
        public override long Position
        {
            get { return input.Position; }
            set { input.Position = value; }
        }
        /// <summary>
        /// Closes the current stream and releases any resources (such as sockets and file handles) associated with the current stream.
        /// </summary>
        public override void Close()
        {
            if (input == null)
                return;
            input.Close();
            input = null;
            buf = null;
        }
        /// <summary>
        /// When overridden in a derived class, clears all buffers for this stream and causes any buffered data to be written to the underlying device.
        /// </summary>
        /// <exception cref="T:System.IO.IOException">
        /// An I/O error occurs.
        /// </exception>
        public override void Flush()
        {
            input.Flush();
        }
        /// <summary>
        /// When overridden in a derived class, sets the position within the current stream.
        /// </summary>
        /// <param name="offset">A byte offset relative to the <paramref name="origin"/> parameter.</param>
        /// <param name="origin">A value of type <see cref="T:System.IO.SeekOrigin"/> indicating the reference point used to obtain the new position.</param>
        /// <returns>
        /// The new position within the current stream.
        /// </returns>
        /// <exception cref="T:System.IO.IOException">
        /// An I/O error occurs.
        /// </exception>
        /// <exception cref="T:System.NotSupportedException">
        /// The stream does not support seeking, such as if the stream is constructed from a pipe or console output.
        /// </exception>
        /// <exception cref="T:System.ObjectDisposedException">
        /// Methods were called after the stream was closed.
        /// </exception>
        public override long Seek(long offset, SeekOrigin origin)
        {
            return input.Seek(offset, origin);
        }
        /// <summary>
        /// When overridden in a derived class, sets the length of the current stream.
        /// </summary>
        /// <param name="value">The desired length of the current stream in bytes.</param>
        /// <exception cref="T:System.IO.IOException">
        /// An I/O error occurs.
        /// </exception>
        /// <exception cref="T:System.NotSupportedException">
        /// The stream does not support both writing and seeking, such as if the stream is constructed from a pipe or console output.
        /// </exception>
        /// <exception cref="T:System.ObjectDisposedException">
        /// Methods were called after the stream was closed.
        /// </exception>
        public override void SetLength(long value)
        {
            input.SetLength(value);
        }
        /// <summary>
        /// When overridden in a derived class, writes a sequence of bytes to the current stream and advances the current position within this stream by the number of bytes written.
        /// </summary>
        /// <param name="buffer">An array of bytes. This method copies <paramref name="count"/> bytes from <paramref name="buffer"/> to the current stream.</param>
        /// <param name="offset">The zero-based byte offset in <paramref name="buffer"/> at which to begin copying bytes to the current stream.</param>
        /// <param name="count">The number of bytes to be written to the current stream.</param>
        /// <exception cref="T:System.ArgumentException">
        /// The sum of <paramref name="offset"/> and <paramref name="count"/> is greater than the buffer length.
        /// </exception>
        /// <exception cref="T:System.ArgumentNullException">
        /// 	<paramref name="buffer"/> is null.
        /// </exception>
        /// <exception cref="T:System.ArgumentOutOfRangeException">
        /// 	<paramref name="offset"/> or <paramref name="count"/> is negative.
        /// </exception>
        /// <exception cref="T:System.IO.IOException">
        /// An I/O error occurs.
        /// </exception>
        /// <exception cref="T:System.NotSupportedException">
        /// The stream does not support writing.
        /// </exception>
        /// <exception cref="T:System.ObjectDisposedException">
        /// Methods were called after the stream was closed.
        /// </exception>
        public override void Write(byte[] buffer, int offset, int count)
        {
            input.Write(buffer, offset, count);
        }
        /// <summary>
        /// Writes a byte to the current position in the stream and advances the position within the stream by one byte.
        /// </summary>
        /// <param name="value">The byte to write to the stream.</param>
        /// <exception cref="T:System.IO.IOException">
        /// An I/O error occurs.
        /// </exception>
        /// <exception cref="T:System.NotSupportedException">
        /// The stream does not support writing, or the stream is already closed.
        /// </exception>
        /// <exception cref="T:System.ObjectDisposedException">
        /// Methods were called after the stream was closed.
        /// </exception>
        public override void WriteByte(byte value)
        {
            input.WriteByte(value);
        }
	}
}
