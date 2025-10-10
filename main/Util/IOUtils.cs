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
    using NPOI.SS.UserModel;
    using System;
    using System.IO;
    using System.Runtime.Serialization;

    public class IOUtils
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(IOUtils));

        /// <summary>
        /// The default buffer size to use for the skip() methods.
        /// </summary>
        private static int SKIP_BUFFER_SIZE = 2048;
        private static byte[] SKIP_BYTE_BUFFER;
        /// <summary>
        /// The current set global allocation limit override,
        /// -1 means limits are applied per record type. 
        /// The current set global allocation limit override,
        /// </summary>
        private static int BYTE_ARRAY_MAX_OVERRIDE = -1;

        /// <summary>
        /// The max init size of ByteArrayOutputStream.
        /// -1 means init size of ByteArrayOutputStream could be up to Integer.MAX_VALUE
        /// </summary>
        private static int MAX_BYTE_ARRAY_INIT_SIZE = -1;

        /// <summary>
        /// The default size of the bytearray used while reading input streams. This is meant to be pretty small.
        /// </summary>
        private static readonly int DEFAULT_BUFFER_SIZE = 4096;

        /**
         * If this value is set to > 0, {@link #safelyAllocate(long, int)} will ignore the
         * maximum record length parameter.  This is designed to allow users to bypass
         * the hard-coded maximum record lengths if they are willing to accept the risk
         * of an OutOfMemoryException.
         *
         * @param maxOverride
         * @since 4.0.0
         */
        public static void SetByteArrayMaxOverride(int maxOverride)
        {
            BYTE_ARRAY_MAX_OVERRIDE = maxOverride;
        }
        /// <summary>
        /// Peeks at the first 8 bytes of the stream. Returns those bytes, but
        ///  with the stream unaffected. Requires a stream that supports mark/reset,
        ///  or a PushbackInputStream. If the stream has &gt;0 but &lt;8 bytes, 
        ///  remaining bytes will be zero.
        /// @throws EmptyFileException if the stream is empty
        /// </summary>
        public static byte[] PeekFirst8Bytes(InputStream stream)
        {
            return PeekFirstNBytes(stream, 8);
            // We want to peek at the first 8 bytes
            //stream.Mark(8);

            //byte[] header = new byte[8];
            //int read = IOUtils.ReadFully(stream, header);

            //if(read < 1)
            //    throw new EmptyFileException();

            //// Wind back those 8 bytes
            //if(stream is PushbackInputStream)
            //{
            //    //PushbackInputStream pin = (PushbackInputStream)stream;
            //    //pin.Unread(header, 0, read);
            //    stream.Position -= read;
            //}
            //else
            //{
            //    stream.Reset();
            //}

            //return header;
        }

        public static byte[] PeekFirstNBytes(Stream stream, int limit)
        {
            long mark =  stream.Position;
            if(stream is InputStream is1)
            {
                is1.Mark(limit);
            }
            ByteArrayOutputStream bos = new ByteArrayOutputStream(limit);
            if(stream is ByteArrayInputStream inputStream)
                Copy(new BoundedInputStream(inputStream, limit), bos);
            else
            {
                //MemoryStream ms = new MemoryStream();
                byte[] buffer = new byte[limit];
                stream.Read(buffer, 0, limit);
                Copy(new BoundedInputStream(new ByteArrayInputStream(buffer), limit), bos);
            }

            int readBytes = (int)bos.Length;
            if(readBytes == 0)
            {
                throw new EmptyFileException();
            }
            if(readBytes < limit)
            {
                bos.Write(new byte[limit - readBytes]);
            }
            byte[] peekedBytes = bos.ToByteArray();
            if(stream is PushbackInputStream)
            {
                //PushbackInputStream pin = (PushbackInputStream)stream;
                //pin.unread(peekedBytes, 0, readBytes);
                stream.Position -= peekedBytes.Length;
            }
            else if(stream is InputStream is2)
            {
                is2.Reset();
            }
            else
            {
                stream.Position = mark;
            }

            return peekedBytes;
        }

        /// <summary>
        /// Reads all the data from the input stream, and returns
        /// the bytes Read.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        /// <remarks>Tony Qu changed the code</remarks>
        public static byte[] ToByteArray(Stream stream)
        {
            return ToByteArray(stream, Int32.MaxValue);
        }

        /// <summary>
        /// Reads up to {@code length} bytes from the input stream, and returns the bytes read.
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public static byte[] ToByteArray(Stream stream, int length)
        {
            using (ByteArrayOutputStream baos = new ByteArrayOutputStream(length == Int32.MaxValue ? 4096 : length))
            {
                byte[] buffer = new byte[4096];
                int totalBytes = 0, readBytes;
                do
                {
                    readBytes = stream.Read(buffer, 0, Math.Min(buffer.Length, length - totalBytes));
                    totalBytes += Math.Max(readBytes, 0);
                    if(readBytes > 0)
                    {
                        baos.Write(buffer, 0, readBytes);
                    }
                } while(totalBytes < length && readBytes > 0);
            
                if(length != Int32.MaxValue && totalBytes < length)
                {
                    throw new IOException("unexpected EOF");
                }
            
                return baos.ToByteArray();
            }
        }

        public static byte[] ToByteArray(ByteBuffer buffer, int length)
        {
            if (buffer.HasBuffer && buffer.Offset == 0)
            {
                // The backing array should work out fine for us
                return buffer.Buffer;
            }

            byte[] data = new byte[length];
            buffer.Read(data);
            return data;
        }


        /// <summary>
        /// Reads the fully.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="b">The b.</param>
        /// <returns></returns>
        public static int ReadFully(Stream stream, byte[] b)
        {
            return ReadFully(stream, b, 0, b.Length);
        }

        /// <summary>
        /// Same as the normal InputStream#read(byte[], int, int), but tries to ensure
        /// that the entire len number of bytes is read
        /// 
        /// If the end of file is reached before any bytes are read, returns -1.
        /// If the end of the file is reached after some bytes are read, returns the
        /// number of bytes read. If the end of the file isn't reached before the
        /// buffer has no more remaining capacity, will return len bytes
        /// </summary>
        /// <param name="stream">the stream from which the data is read.</param>
        /// <param name="b">the buffer into which the data is read.</param>
        /// <param name="off">the start offset in array b at which the data is written.</param>
        /// <param name="len">the maximum number of bytes to read.</param>
        /// <returns>the number of bytes read or -1 if no bytes were read</returns>
        public static int ReadFully(Stream stream, byte[] b, int off, int len)
        {
            int total = 0;
            while (true)
            {
                int got = stream.Read(b, off + total, len - total - off);
                if (got <= 0)
                {
                    return (total == 0) ? -1 : total;
                }
                total += got;
                if (total == len)
                {
                    return total;
                }
            }
        }

        /// <summary>
        /// <para>
        /// Write a POI Document (<see cref="NPOI.SS.UserModel.Workbook}, {@link NPOI.sl.UserModel.SlideShow" />, etc) to an output stream and close the output stream.
        /// This will attempt to close the output stream at the end even if there was a problem writing the document to the stream.
        /// </para>
        /// <para>
        /// If you are using Java 7 or higher, you may prefer to use a try-with-resources statement instead.
        /// This function exists for Java 6 code.
        /// </para>
        /// </summary>
        /// <param name="doc"> a writeable document to write to the output stream</param>
        /// <param name="out"> the output stream that the document is written to</param>
        /// <exception cref="IOException">IOException</exception>
        public static void Write(POIDocument doc, OutputStream out1)
        {

            try
            {
                doc.Write(out1);
            }
            finally
            {
                CloseQuietly(out1);
            }
        }

        public static void Write(IWorkbook doc, OutputStream out1)
        {
            try
            {
                doc.Write(out1);
            }
            finally
            {
                CloseQuietly(out1);
            }
        }

        /// <summary>
        /// <para>
        /// Write a POI Document (<see cref="IWorkbook" />, etc) to an output stream and close the output stream.
        /// This will attempt to close the output stream at the end even if there was a problem writing the document to the stream.
        /// This will also attempt to close the document, even if an error occurred while writing the document or closing the output stream.
        /// </para>
        /// <para>
        /// If you are using Java 7 or higher, you may prefer to use a try-with-resources statement instead.
        /// This function exists for Java 6 code.
        /// </para>
        /// </summary>
        /// <param name="doc"> a writeable and closeable document to write to the output stream, then close</param>
        /// <param name="out"> the output stream that the document is written to</param>
        /// <exception cref="IOException">IOException</exception>
        public static void WriteAndClose(POIDocument doc, OutputStream out1)
        {

            try
            {
                Write(doc, out1);
            }
            finally
            {
                CloseQuietly(doc);
            }
        }

        /// <summary>
        /// <para>
        /// Like <see cref="WriteAndClose(POIDocument, OutputStream)" />, but for writing to a File instead of an OutputStream.
        /// This will attempt to close the document, even if an error occurred while writing the document.
        /// </para>
        /// <para>
        /// If you are using Java 7 or higher, you may prefer to use a try-with-resources statement instead.
        /// This function exists for Java 6 code.
        /// </para>
        /// </summary>
        /// <param name="doc"> a writeable and closeable document to write to the output file, then close</param>
        /// <param name="out"> the output file that the document is written to</param>
        /// <exception cref="IOException">IOException</exception>
        public static void WriteAndClose(POIDocument doc, FileInfo out1)
        {

            try
            {
                doc.Write(out1);
            }
            finally
            {
                CloseQuietly(doc);
            }
        }

        /// <summary>
        /// <para>
        /// Like <see cref="write(POIDocument, File)" />, but for writing a POI Document in place (to the same file that it was opened from).
        /// This will attempt to close the document, even if an error occurred while writing the document.
        /// </para>
        /// <para>
        /// If you are using Java 7 or higher, you may prefer to use a try-with-resources statement instead.
        /// This function exists for Java 6 code.
        /// </para>
        /// </summary>
        /// <param name="doc"> a writeable document to write in-place</param>
        /// <exception cref="IOException">IOException</exception>
        public static void WriteAndClose(POIDocument doc)
        {

            try
            {
                doc.Write();
            }
            finally
            {
                CloseQuietly(doc);
            }
        }

        public static void WriteAndClose(IWorkbook doc, OutputStream out1)
        {

            try
            {
                doc.Write(out1);
            }
            finally
            {
                CloseQuietly(doc);
            }
        }

        /// <summary>
        /// Copies all the data from the given InputStream to the OutputStream. It
        /// leaves both streams open, so you will still need to close them once done.
        /// </summary>
        /// <param name="inp"></param>
        /// <param name="out1"></param>
        public static void Copy(Stream inp, Stream out1)
        {
            byte[] buff = new byte[4096];
            //inp.Position = 0;
            int count;
            while ((count = inp.Read(buff, 0, buff.Length)) >0)
            {
                if (count < -1)
                {
                    throw new RecordFormatException("Can't have read < -1 bytes");
                }
                out1.Write(buff, 0, count);
            }
        }

        public static long CalculateChecksum(byte[] data)
        {
            CRC32 sum = new CRC32();
            return (long)sum.ByteCRC(ref data);
        }

        ///<summary>
        ///Quietly (no exceptions) close Closable resource. In case of error it will
        ///be printed to {@link IOUtils} class logger.
        ///</summary>
        ///<param name="closeable">resource to close</param>
        public static void CloseQuietly(Stream closeable )
        {
            // no need to log a NullPointerException here
            if (closeable == null)
            {
                return;
            }
            try
            {
                closeable.Close();
            }
            catch (Exception exc)
            {
                logger.Log(POILogger.ERROR, "Unable to close resource: " + exc, exc);
            }
        }

        public static void CloseQuietly(ICloseable closeable)
        {
            // no need to log a NullPointerException here
            if (closeable == null)
            {
                return;
            }
            try
            {
                closeable.Close();
            }
            catch (Exception exc)
            {
                logger.Log(POILogger.ERROR, "Unable to close resource: " + exc, exc);
            }
        }

        /// <summary>
        /// <para>
        /// Skips bytes from an input byte stream.
        /// This implementation guarantees that it will read as many bytes
        /// as possible before giving up; this may not always be the case for
        /// skip() implementations in subclasses of <see cref="InputStream"/>.
        /// </para>
        /// <para>
        /// Note that the implementation uses <see cref="InputStream.Read(byte[], int, int)" /> rather
        /// than delegating to <see cref="InputStream.Skip(long)" />.
        /// This means that the method may be considerably less efficient than using the actual skip implementation,
        /// this is done to guarantee that the correct number of bytes are skipped.
        /// </para>
        /// <para>
        /// </para>
        /// <para>
        /// This mimics POI's <see cref="ReadFully(InputStream, byte[])" />.
        /// If the end of file is reached before any bytes are read, returns <tt>-1</tt>. If
        /// the end of the file is reached After some bytes are read, returns the
        /// number of bytes read. If the end of the file isn't reached before <tt>len</tt>
        /// bytes have been read, will return <tt>len</tt> bytes.
        /// </para>
        /// <para>
        /// </para>
        /// <para>
        /// Copied nearly verbatim from commons-io 41a3e9c
        /// </para>
        /// </summary>
        /// <param name="input">byte stream to skip</param>
        /// <param name="toSkip">number of bytes to skip.</param>
        /// <returns>number of bytes actually skipped.</returns>
        /// <exception cref="IOException">             if there is a problem reading the file</exception>
        /// <exception cref="ArgumentException">if toSkip is negative</exception>
        /// @see InputStream#skip(long)
        ///
        public static long SkipFully(InputStream input, long toSkip)
        {
            if (toSkip < 0) {
                throw new ArgumentException("Skip count must be non-negative, actual: " + toSkip);
            }
            if (toSkip == 0) {
                return 0L;
            }
            /*
             * N.B. no need to synchronize this because: - we don't care if the buffer is created multiple times (the data
             * is ignored) - we always use the same size buffer, so if it it is recreated it will still be OK (if the buffer
             * size were variable, we would need to synch. to ensure some other thread did not create a smaller one)
             */
            if (SKIP_BYTE_BUFFER == null) {
                SKIP_BYTE_BUFFER = new byte[SKIP_BUFFER_SIZE];
            }
            long remain = toSkip;
            while (remain > 0) {
                // See https://issues.apache.org/jira/browse/IO-203 for why we use read() rather than delegating to skip()
                long n = input.Read(SKIP_BYTE_BUFFER, 0, (int) Math.Min(remain, SKIP_BUFFER_SIZE));
                if (n <= 0) { // EOF
                    break;
                }
                remain -= n;
            }
            if (toSkip == remain) {
                return -1L;
            }
            return toSkip - remain;
        }

        public static byte[] SafelyAllocate(long length, int maxLength)
        {
            SafelyAllocateCheck(length, maxLength);

            //CheckByteSizeLimit((int)length);
            if(BYTE_ARRAY_MAX_OVERRIDE > 0)
            {
                if(length > BYTE_ARRAY_MAX_OVERRIDE)
                {
                    ThrowRFE(length, BYTE_ARRAY_MAX_OVERRIDE);
                }
            }
            else if(length > maxLength)
            {
                ThrowRFE(length, maxLength);
            }
            return new byte[(int)length];
        }

        public static void SafelyAllocateCheck(long length, int maxLength)
        {
            if (length < 0L)
            {
                throw new RecordFormatException("Can't allocate an array of length < 0, but had " + length + " and " + maxLength);
            }
            if (length > (long)int.MaxValue)
            {
                throw new RecordFormatException("Can't allocate an array > " + int.MaxValue);
            }
            CheckLength(length, maxLength);
        }

        private static void CheckLength(long length, int maxLength)
        {
            if (BYTE_ARRAY_MAX_OVERRIDE > 0)
            {
                if (length > BYTE_ARRAY_MAX_OVERRIDE)
                {
                    ThrowRFE(length, BYTE_ARRAY_MAX_OVERRIDE);
                }
            }
            else if (length > maxLength)
            {
                ThrowRFE(length, maxLength);
            }
        }

        private static void ThrowRFE(long length, int maxLength)
        {
            throw new RecordFormatException(string.Format("Tried to allocate an array of length {0}" +
                            ", but the maximum length for this record type is {1}.\n" +
                            "If the file is not corrupt and not large, please open an issue on GitHub to request \n" +
                            "increasing the maximum allowable size for this record type.\n" +
                            "You can set a higher override value with IOUtils.SetByteArrayMaxOverride()", length, maxLength));
        }

        private static void ThrowRecordTruncationException(int maxLength)
        {
            throw new RecordFormatException(string.Format("Tried to read data but the maximum length " +
                    "for this record type is {0}.\n" +
                    "If the file is not corrupt and not large, please open an issue on GitHub to request \n" +
                    "increasing the maximum allowable size for this record type.\n" +
                    "You can set a higher override value with IOUtils.SetByteArrayMaxOverride()", maxLength));
        }
    }
}
