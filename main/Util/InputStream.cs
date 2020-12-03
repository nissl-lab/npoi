using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.Util
{
    public abstract class InputStream : Stream
    {
        // MAX_SKIP_BUFFER_SIZE is used to determine the maximum buffer size to
        // use when skipping.
        private static int MAX_SKIP_BUFFER_SIZE = 2048;
        /// <summary>
        /// Reads the next byte of data from the input stream. The value byte is
        /// returned as an <c>int</c> in the range <c>0</c> to
        /// <c>255</c>. If no byte is available because the end of the stream
        /// has been reached, the value <c>-1</c> is returned. This method
        /// blocks until input data is available, the end of the stream is detected,
        /// or an exception is thrown.
        /// 
        /// A subclass must provide an implementation of this method.
        /// </summary>
        /// <returns>
        /// the next byte of data, or <c>-1</c> if the end of the
        /// stream is reached.
        /// </returns>
        /// <exception cref="IOException">if an I/O error occurs</exception>
        public abstract int Read();
        /// <summary>
        /// Reads some number of bytes from the input stream and stores them into
        /// the buffer array <c>b</c>. The number of bytes actually read is
        /// returned as an integer.  This method blocks until input data is
        /// available, end of file is detected, or an exception is thrown.
        ///
        /// <p> If the length of <c>b</c> is zero, then no bytes are read and
        /// <c>0</c> is returned; otherwise, there is an attempt to read at
        /// least one byte. If no byte is available because the stream is at the
        /// end of the file, the value <c>-1</c> is returned; otherwise, at
        /// least one byte is read and stored into <c>b</c>.</p>
        ///
        /// <p> The first byte read is stored into element <c>b[0]</c>, the
        /// next one into <c>b[1]</c>, and so on. The number of bytes read is,
        /// at most, equal to the length of <c>b</c>. Let <i>k</i> be the
        /// number of bytes actually read; these bytes will be stored in elements
        /// <c>b[0]</c> through <c>b[</c><i>k</i><c>-1]</c>,
        /// leaving elements <c>b[</c><i>k</i><c>]</c> through
        /// <c>b[b.length-1]</c> unaffected.</p>
        ///
        /// <p> The <c>read(b)</c> method for class <c>InputStream</c>
        /// has the same effect as: <pre><c> read(b, 0, b.length) </c></pre></p>
        /// </summary>
        /// <param name="b">the buffer into which the data is read.</param>
        /// <returns>
        /// the total number of bytes read into the buffer, or
        /// <c>-1</c> if there is no more data because the end of
        /// the stream has been reached.
        /// </returns>
        /// <exception cref="IOException">If the first byte cannot be read for any reason
        /// other than the end of the file, if the input stream has been closed, or
        /// if some other I/O error occurs.</exception>
        /// <exception cref="NullReferenceException">if <c>b</c> is <c>null</c>.</exception>
        /// <see cref="Read(byte[], int, int)"/>
        public virtual int Read(byte[] b)
        {
            return Read(b, 0, b.Length);
        }
        /// <summary>
        /// Reads up to <c>len</c> bytes of data from the input stream into
        /// an array of bytes.  An attempt is made to read as many as
        /// <c>len</c> bytes, but a smaller number may be read.
        /// The number of bytes actually read is returned as an integer.
        ///
        /// <p> This method blocks until input data is available, end of file is
        /// detected, or an exception is thrown.</p>
        ///
        /// <p> If <c>len</c> is zero, then no bytes are read and
        /// <c>0</c> is returned; otherwise, there is an attempt to read at
        /// least one byte. If no byte is available because the stream is at end of
        /// file, the value <c>-1</c> is returned; otherwise, at least one
        /// byte is read and stored into <c>b</c>.</p>
        ///
        /// <p> The first byte read is stored into element <c>b[off]</c>, the
        /// next one into <c>b[off+1]</c>, and so on. The number of bytes read
        /// is, at most, equal to <c>len</c>. Let <i>k</i> be the number of
        /// bytes actually read; these bytes will be stored in elements
        /// <c>b[off]</c> through <c>b[off+</c><i>k</i><c>-1]</c>,
        /// leaving elements <c>b[off+</c><i>k</i><c>]</c> through
        /// <c>b[off+len-1]</c> unaffected.</p>
        ///
        /// <p> In every case, elements <c>b[0]</c> through
        /// <c>b[off]</c> and elements <c>b[off+len]</c> through
        /// <c>b[b.length-1]</c> are unaffected.</p>
        ///
        /// <p> The <c>read(b,</c> <c>off,</c> <c>len)</c> method
        /// for class <c>InputStream</c> simply calls the method
        /// <c>read()</c> repeatedly. If the first such call results in an
        /// <c>IOException</c>, that exception is returned from the call to
        /// the <c>read(b,</c> <c>off,</c> <c>len)</c> method.  If
        /// any subsequent call to <c>read()</c> results in a
        /// <c>IOException</c>, the exception is caught and treated as if it
        /// were end of file; the bytes read up to that point are stored into
        /// <c>b</c> and the number of bytes read before the exception
        /// occurred is returned. The default implementation of this method blocks
        /// until the requested amount of input data <c>len</c> has been read,
        /// end of file is detected, or an exception is thrown. Subclasses are encouraged
        /// to provide a more efficient implementation of this method.</p>
        /// </summary>
        /// <param name="b">the buffer into which the data is read.</param>
        /// <param name="off">the start offset in array <c>b</c> at which the data is written.</param>
        /// <param name="len">the maximum number of bytes to read.</param>
        /// <returns>
        /// the total number of bytes read into the buffer, or
        /// <c>-1</c> if there is no more data because the end of
        /// the stream has been reached.</returns>
        /// <exception cref="IOException">If the first byte cannot be read for any reason
        /// other than end of file, or if the input stream has been closed, or if
        /// some other I/O error occurs.</exception>
        /// <exception cref="NullReferenceException">If <c>b</c> is <c>null</c>.</exception>
        /// <exception cref="IndexOutOfRangeException">If <c>off</c> is negative,
        /// <c>len</c> is negative, or <c>len</c> is greater than
        /// <c>b.length - off</c></exception>
        /// <see cref="Read()"/>
        public override int Read(byte[] b, int off, int len)
        {
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

            int c = Read();
            if (c == -1)
            {
                return -1;
            }
            b[off] = (byte)c;

            int i = 1;
            try
            {
                for (; i < len; i++)
                {
                    c = Read();
                    if (c == -1)
                    {
                        break;
                    }
                    b[off + i] = (byte)c;
                }
            }
            catch (IOException ee)
            {
            }
            return i;
        }
        /// <summary>
        /// Skips over and discards <c>n</c> bytes of data from this input
        /// stream. The <c>skip</c> method may, for a variety of reasons, end
        /// up skipping over some smaller number of bytes, possibly <c>0</c>.
        /// This may result from any of a number of conditions; reaching end of file
        /// before <c>n</c> bytes have been skipped is only one possibility.
        /// The actual number of bytes skipped is returned. If {@code n} is
        /// negative, the {@code skip} method for class {@code InputStream} always
        /// returns 0, and no bytes are skipped. Subclasses may handle the negative
        /// value differently.
        ///
        /// <p> The <c>skip</c> method of this class creates a
        /// byte array and then repeatedly reads into it until <c>n</c> bytes
        /// have been read or the end of the stream has been reached. Subclasses are
        /// encouraged to provide a more efficient implementation of this method.
        /// For instance, the implementation may depend on the ability to seek.</p>
        /// </summary>
        /// <param name="n">the number of bytes to be skipped.</param>
        /// <returns>the actual number of bytes skipped.</returns>
        /// <exception cref="IOException">if the stream does not support seek,
        /// or if some other I/O error occurs.
        /// </exception>
        public virtual long Skip(long n)
        {
            long remaining = n;
            int nr;

            if (n <= 0)
            {
                return 0;
            }

            int size = (int)Math.Min(MAX_SKIP_BUFFER_SIZE, remaining);
            byte[] skipBuffer = new byte[size];
            while (remaining > 0)
            {
                nr = Read(skipBuffer, 0, (int)Math.Min(size, remaining));
                if (nr < 0)
                {
                    break;
                }
                remaining -= nr;
            }

            return n - remaining;
        }

        /// <summary>
        /// Returns an estimate of the number of bytes that can be read (or
        /// skipped over) from this input stream without blocking by the next
        /// invocation of a method for this input stream. The next invocation
        /// might be the same thread or another thread.  A single read or skip of this
        /// many bytes will not block, but may read or skip fewer bytes.
        ///
        /// <p> Note that while some implementations of {@code InputStream} will return
        /// the total number of bytes in the stream, many will not.  It is
        /// never correct to use the return value of this method to allocate
        /// a buffer intended to hold all data in this stream.</p>
        ///
        /// <p> A subclass' implementation of this method may choose to throw an
        /// {@link IOException} if this input stream has been closed by
        /// invoking the {@link #close()} method.</p>
        ///
        /// <p> The {@code available} method for class {@code InputStream} always
        /// returns {@code 0}.</p>
        ///
        /// <p> This method should be overridden by subclasses.</p>
        /// </summary>
        /// <exception cref="IOException">if an I/O error occurs.</exception>
        public virtual int Available()
        {
            return 0;
        }
        /// <summary>
        /// Closes this input stream and releases any system resources associated
        /// with the stream.
        ///
        /// <p> The <c>Close</c> method of <c>InputStream</c> does nothing.</p>
        /// </summary>
        /// <exception cref="IOException">if an I/O error occurs.</exception>
        public override void Close()
        {

        }
        /// <summary>
        /// Marks the current position in this input stream. A subsequent call to
        /// the <c>reset</c> method repositions this stream at the last marked
        /// position so that subsequent reads re-read the same bytes.
        ///
        /// <p> The <c>readlimit</c> arguments tells this input stream to
        /// allow that many bytes to be read before the mark position gets
        /// invalidated.</p>
        ///
        /// <p> The general contract of <c>mark</c> is that, if the method
        /// <c>markSupported</c> returns <c>true</c>, the stream somehow
        /// remembers all the bytes read after the call to <c>mark</c> and
        /// stands ready to supply those same bytes again if and whenever the method
        /// <c>reset</c> is called.  However, the stream is not required to
        /// remember any data at all if more than <c>readlimit</c> bytes are
        /// read from the stream before <c>reset</c> is called.</p>
        ///
        /// <p> Marking a closed stream should not have any effect on the stream.</p>
        ///
        /// <p> The <c>mark</c> method of <c>InputStream</c> does
        /// nothing.</p>
        /// </summary>
        /// <param name="readlimit">the maximum limit of bytes that can be read before
        /// the mark position becomes invalid.
        /// </param>
        /// <see cref="Reset"/>
        public virtual void Mark(int readlimit) { }
        /// <summary>
        /// Repositions this stream to the position at the time the
        /// <c>mark</c> method was last called on this input stream.
        ///
        /// <p> The general contract of <c>reset</c> is:</p>
        ///
        /// <ul>
        /// <li> If the method <c>markSupported</c> returns
        /// <c>true</c>, then:
        ///
        ///     <ul><li> If the method <c>mark</c> has not been called since
        ///     the stream was created, or the number of bytes read from the stream
        ///     since <c>mark</c> was last called is larger than the argument
        ///     to <c>mark</c> at that last call, then an
        ///     <c>IOException</c> might be thrown.</li>
        ///
        ///     <li> If such an <c>IOException</c> is not thrown, then the
        ///     stream is reset to a state such that all the bytes read since the
        ///     most recent call to <c>mark</c> (or since the start of the
        ///     file, if <c>mark</c> has not been called) will be resupplied
        ///     to subsequent callers of the <c>read</c> method, followed by
        ///     any bytes that otherwise would have been the next input data as of
        ///     the time of the call to <c>reset</c>. </li>
        ///
        /// <li> If the method <c>markSupported</c> returns
        /// <c>false</c>, then:
        ///
        ///     <ul><li> The call to <c>reset</c> may throw an
        ///     <c>IOException</c>.</li>
        ///
        ///     <li> If an <c>IOException</c> is not thrown, then the stream
        ///     is reset to a fixed state that depends on the particular type of the
        ///     input stream and how it was created. The bytes that will be supplied
        ///     to subsequent callers of the <c>read</c> method depend on the
        ///     particular type of the input stream. </li></ul></li></ul></li></ul>
        ///
        /// <p>The method <c>reset</c> for class <c>InputStream</c>
        /// does nothing except throw an <c>IOException</c>.</p>
        /// </summary>
        public virtual void Reset()
        {
            throw new IOException("mark/reset not supported");
        }
        /// <summary>
        /// Tests if this input stream supports the <c>mark</c> and
        /// <c>reset</c> methods. Whether or not <c>mark</c> and
        /// <c>reset</c> are supported is an invariant property of a
        /// particular input stream instance. The <c>markSupported</c> method
        /// of <c>InputStream</c> returns <c>false</c>.
        /// </summary>
        /// <returns>
        /// <c>true</c> if this stream instance supports the mark
        /// and reset methods; <c>false</c> otherwise.
        /// <see cref="InputStream.Mark(int)"/>
        /// <see cref="InputStream.Reset"/>
        /// </returns>
        public virtual bool MarkSupported()
        {
            return false;
        }
    }

    public class FileInputStream : InputStream
    {
        Stream inner;
        public FileInputStream(Stream fs)
        {
            this.inner = fs;
        }

        public override bool CanRead
        {
            get
            {
                return inner.CanRead;
            }
        }
        public override bool CanSeek
        {
            get { return false; }
        } 
        public override bool CanWrite
        {
            get { return false; }
        } 
        public override long Length
        {
            get
            {
                return inner.Length;
            }
        }

        public override long Position
        {
            get { return inner.Position;}
            set { inner.Position = value;}
        }

        public override void Flush()
        {
            throw new NotImplementedException();
        }

        public override int Read()
        {
            return inner.ReadByte();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotImplementedException();
        }

        public override void SetLength(long value)
        {
            throw new NotImplementedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotImplementedException();
        }

        public override void Close()
        {
            if (inner != null)
                inner.Close();
        }
    }
}
