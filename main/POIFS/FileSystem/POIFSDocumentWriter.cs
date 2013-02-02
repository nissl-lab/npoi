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

namespace NPOI.POIFS.FileSystem
{
    /// <summary>
    /// This class provides a wrapper over an OutputStream so that Document
    /// writers can't accidently go over their size limits
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    [Obsolete]
    public class POIFSDocumentWriter:Stream
    {
        private int limit;
        private Stream stream;
        private int written;

        /// <summary>
        /// Create a POIFSDocumentWriter
        /// </summary>
        /// <param name="stream">the OutputStream to which the data is actually</param>
        /// <param name="limit">the maximum number of bytes that can be written</param>
        public POIFSDocumentWriter(Stream stream, int limit)
        {
            this.stream = stream;
            this.limit = limit;
            this.written = 0;
        }
        /// <summary>
        /// Closes this output stream and releases any system resources
        /// associated with this stream. The general contract of close is
        /// that it closes the output stream. A closed stream cannot
        /// perform output operations and cannot be reopened.
        /// </summary>
        public override void Close()
        {
            this.stream.Close();
        }
        /// <summary>
        /// Flushes this output stream and forces any buffered output bytes
        /// to be written out.
        /// </summary>
        public override void Flush()
        {
            this.stream.Flush();
        }

        private void LimitCheck(int toBeWritten)
        {
            if ((this.written + toBeWritten) > this.limit)
            {
                throw new IOException("tried to write too much data");
            }
            this.written += toBeWritten;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotImplementedException();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return 0L;
        }

        public override void SetLength(long value)
        {
            throw new NotImplementedException();
        }


        public void Write(int b)
        {
            LimitCheck(1);
            stream.WriteByte((byte)b);
        }

        /// <summary>
        /// Writes b.length bytes from the specified byte array
        /// to this output stream.
        /// </summary>
        /// <param name="b">the data.</param>
        public void Write(byte[] b)
        {
            this.Write(b, 0, b.Length);
        }
        /// <summary>
        /// Writes len bytes from the specified byte array starting at
        /// offset off to this output stream.  The general contract for
        /// write(b, off, len) is that some of the bytes in the array b are
        /// written to the output stream in order; element b[off] is the
        /// first byte written and b[off+len-1] is the last byte written by
        /// this operation.
        /// If b is null, a NullPointerException is thrown.
        /// If off is negative, or len is negative, or off+len is greater
        /// than the length of the array b, then an
        /// IndexOutOfBoundsException is thrown.
        /// </summary>
        /// <param name="b">the data.</param>
        /// <param name="off">the start offset in the data.</param>
        /// <param name="len">the number of bytes to write.</param>
        public override void Write(byte[] b, int off, int len)
        {
            this.LimitCheck(len);
            this.stream.Write(b, off, len);
        }

        /// <summary>
        /// Writes the specified byte to this output stream. The general
        /// contract for write is that one byte is written to the output
        /// stream. The byte to be written is the eight low-order bits of
        /// the argument b. The 24 high-order bits of b are ignored.
        /// </summary>
        /// <param name="b">the byte.</param>
        public override void WriteByte(byte b)
        {
            this.LimitCheck(1);
            this.stream.WriteByte(b);
        }
        /// <summary>
        /// write the rest of the document's data (fill in at the end)
        /// </summary>
        /// <param name="totalLimit">the actual number of bytes the corresponding         
        /// document must fill</param>
        /// <param name="fill">the byte to fill remaining space with</param>
        public virtual void WriteFiller(int totalLimit, byte fill)
        {
            if (totalLimit > this.written)
            {
                byte[] buffer = new byte[totalLimit - this.written];
                for (int i = 0; i < buffer.Length; i++)
                {
                    buffer[i] = fill;
                }
                this.stream.Write(buffer, 0, buffer.Length);
            }
        }

        /// <summary>
        /// When overridden in a derived class, gets a value indicating whether the current stream supports reading.
        /// </summary>
        /// <value></value>
        /// <returns>true if the stream supports reading; otherwise, false.
        /// </returns>
        public override bool CanRead
        {
            get
            {
                return false;
            }
        }

        /// <summary>
        /// When overridden in a derived class, gets a value indicating whether the current stream supports seeking.
        /// </summary>
        /// <value></value>
        /// <returns>true if the stream supports seeking; otherwise, false.
        /// </returns>
        public override bool CanSeek
        {
            get
            {
                return false;
            }
        }

        /// <summary>
        /// When overridden in a derived class, gets a value indicating whether the current stream supports writing.
        /// </summary>
        /// <value></value>
        /// <returns>true if the stream supports writing; otherwise, false.
        /// </returns>
        public override bool CanWrite
        {
            get
            {
                return true;
            }
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
            get
            {
                return this.stream.Length;
            }
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
            get
            {
                return this.stream.Position;
            }
            set
            {
                this.stream.Position = value;
            }
        }

    }
}
