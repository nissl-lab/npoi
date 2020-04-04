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
    using System;
    using System.IO;
    using System.Runtime.Serialization;

    public class IOUtils
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(IOUtils));
        /**
         * Peeks at the first 8 bytes of the stream. Returns those bytes, but
         *  with the stream unaffected. Requires a stream that supports mark/reset,
         *  or a PushbackInputStream. If the stream has &gt;0 but &lt;8 bytes, 
         *  remaining bytes will be zero.
         * @throws EmptyFileException if the stream is empty
         */
        public static byte[] PeekFirst8Bytes(InputStream stream)
        {
            // We want to peek at the first 8 bytes
            stream.Mark(8);

            byte[] header = new byte[8];
            int read = IOUtils.ReadFully(stream, header);

            if (read < 1)
                throw new EmptyFileException();

            // Wind back those 8 bytes
            if (stream is PushbackInputStream) {
                //PushbackInputStream pin = (PushbackInputStream)stream;
                //pin.Unread(header, 0, read);
                stream.Position -= read;
            } else {
                stream.Reset();
            }

            return header;
        }
        public static byte[] PeekFirst8Bytes(Stream stream)
        {
            // We want to peek at the first 8 bytes
            long mark =  stream.Position;

            byte[] header = new byte[8];
            int read = IOUtils.ReadFully(stream, header);

            if (read < 1)
                throw new EmptyFileException();

            // Wind back those 8 bytes
            if (stream is PushbackInputStream)
            {
                //PushbackInputStream pin = (PushbackInputStream)stream;
                //pin.Unread(header, 0, read);
                stream.Position -= read;
            }
            else
            {
                stream.Position = mark;
            }

            return header;
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
            ByteArrayOutputStream baos = new ByteArrayOutputStream(length == Int32.MaxValue ? 4096 : length);

            byte[] buffer = new byte[4096];
            int totalBytes = 0, readBytes;
            do
            {
                readBytes = stream.Read(buffer, 0, Math.Min(buffer.Length, length - totalBytes));
                totalBytes += Math.Max(readBytes, 0);
                if (readBytes > 0)
                {
                    baos.Write(buffer, 0, readBytes);
                }
            } while (totalBytes < length && readBytes > 0);

            if (length != Int32.MaxValue && totalBytes < length)
            {
                throw new IOException("unexpected EOF");
            }

            return baos.ToByteArray();
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
        /// that the buffer is filled completely if possible, i.e. b.remaining()
        /// returns 0.
        /// If the end of file is reached before any bytes are Read, returns -1.
        /// If the end of the file is reached after some bytes are read, returns the
        /// number of bytes read. If the end of the file isn't reached before the
        /// buffer has no more remaining capacity, will return the number of bytes
        /// that were read.
        /// </summary>
        /// <param name="stream">the stream from which the data is read.</param>
        /// <param name="b">the buffer into which the data is read.</param>
        /// <param name="off">the start offset in array b at which the data is written.</param>
        /// <param name="len">the maximum number of bytes to read.</param>
        /// <returns></returns>

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
                out1.Write(buff, 0, count);
            }
        }

        public static long CalculateChecksum(byte[] data)
        {
            CRC32 sum = new CRC32();
            return (long)sum.ByteCRC(ref data);
        }

        /**
         * Quietly (no exceptions) close Closable resource. In case of error it will
         * be printed to {@link IOUtils} class logger.
         * 
         * @param closeable
         *            resource to close
         */
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
    }

    [Serializable]
    public class EmptyFileException : IOException
    {
        public EmptyFileException()
        {
        }

        public EmptyFileException(string message) : base(message)
        {
        }

        public EmptyFileException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected EmptyFileException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
