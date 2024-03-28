using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Zip.Compression.Streams;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace NPOI.OpenXml4Net.Util
{
    public class ZipSecureFile : ZipFile
    {
        private static double MIN_INFLATE_RATIO = 0.01d;
        private static long MAX_ENTRY_SIZE = 0xFFFFFFFFL;
        // The default maximum size of extracted text 
        private static long MAX_TEXT_SIZE = 10*1024*1024L;
        /// <summary>
        /// Sets the ratio between de- and inflated bytes to detect zipbomb.
        /// It defaults to 1% (= 0.01d), i.e. when the compression is better than
        /// 1% for any given read package part, the parsing will fail indicating a
        /// Zip-Bomb.
        /// </summary>
        /// <param name="ratio">the ratio between de- and inflated bytes to detect zipbomb</param>
        public static void SetMinInflateRatio(double ratio)
        {
            MIN_INFLATE_RATIO = ratio;
        }

        /// <summary>
        /// <para>
        /// Returns the current minimum compression rate that is used.
        /// </para>
        /// <para>
        /// See setMinInflateRatio() for details.
        /// </para>
        /// </summary>
        /// <returns>The min accepted compression-ratio.</returns>
        public static double GetMinInflateRatio()
        {
            return MIN_INFLATE_RATIO;
        }

        /// <summary>
        /// <para>
        /// Sets the maximum file size of a single zip entry. It defaults to 4GB,
        /// i.e. the 32-bit zip format maximum.
        /// </para>
        /// <para>
        /// This can be used to limit memory consumption and protect against
        /// security vulnerabilities when documents are provided by users.
        /// </para>
        /// </summary>
        /// <param name="maxEntrySize">the max. file size of a single zip entry</param>
        public static void SetMaxEntrySize(long maxEntrySize)
        {
            if(maxEntrySize < 0 || maxEntrySize > 0xFFFFFFFFL)
            {
                throw new ArgumentException("Max entry size is bounded [0-4GB].");
            }
            MAX_ENTRY_SIZE = maxEntrySize;
        }

        /// <summary>
        /// <para>
        /// Returns the current maximum allowed uncompressed file size.
        /// </para>
        /// <para>
        /// See setMaxEntrySize() for details.
        /// </para>
        /// </summary>
        /// <returns>The max accepted uncompressed file size.</returns>
        public static long GetMaxEntrySize()
        {
            return MAX_ENTRY_SIZE;
        }

        /// <summary>
        /// <para>
        /// Sets the maximum number of characters of text that are
        /// extracted before an exception is thrown during extracting
        /// text from documents.
        /// </para>
        /// <para>
        /// This can be used to limit memory consumption and protect against
        /// security vulnerabilities when documents are provided by users.
        /// </para>
        /// </summary>
        /// <param name="maxTextSize">the max. file size of a single zip entry</param>
        public static void SetMaxTextSize(long maxTextSize)
        {
            if(maxTextSize < 0 || maxTextSize > 0xFFFFFFFFL)
            {     // don't use MAX_ENTRY_SIZE here!
                throw new ArgumentException("Max text size is bounded [0-4GB], but had " + maxTextSize);
            }
            MAX_TEXT_SIZE = maxTextSize;
        }

        /// <summary>
        /// <para>
        /// Returns the current maximum allowed text size.
        /// </para>
        /// <para>
        /// See setMaxTextSize() for details.
        /// </para>
        /// </summary>
        /// <returns>The max accepted text size.</returns>
        public static long GetMaxTextSize()
        {
            return MAX_TEXT_SIZE;
        }

        public ZipSecureFile(FileStream file, int mode)
            : base(file)
        {

        }

        public ZipSecureFile(FileStream file)
            : base(file)
        {

        }

        public ZipSecureFile(String name)
                : base(name)
        {

        }

        /// <summary>
        /// <para>
        /// Returns an input stream for reading the contents of the specified
        /// zip file entry.
        /// </para>
        /// <para>
        /// Closing this ZIP file will, in turn, close all input
        /// streams that have been returned by invocations of this method.
        /// </para>
        /// </summary>
        /// <param name="entry">the zip file entry</param>
        /// <returns>the input stream for reading the contents of the specified
        /// zip file entry.
        /// </returns>
        /// <exception cref="ZipException">if a ZIP format error has occurred</exception>
        /// <exception cref="IOException">if an I/O error has occurred</exception>
        /// <exception cref="IllegalStateException">if the zip file has been closed</exception>
        public new Stream GetInputStream(ZipEntry entry)
        {
            Stream zipIS = base.GetInputStream(entry);
            return AddThreshold(zipIS);
        }

        public static ThresholdInputStream AddThreshold(Stream zipIS)
        {

            ThresholdInputStream newInner = null;
            if(zipIS is InflaterInputStream)
            {
                //replace inner stream of zipIS by using a ThresholdInputStream instance??
                try
                {
                    FieldInfo f = typeof(FilterInputStream).GetField("in");
                    //f.SetAccessible(true);
                    //InputStream oldInner = (InputStream)f.Get(zipIS);
                    //newInner = new ThresholdInputStream(oldInner, null);
                    //f.Set(zipIS, newInner);
                }
                catch(Exception)
                {
                    //logger.Log(POILogger.WARN, "SecurityManager doesn't allow manipulation via reflection for zipbomb detection - continue with original input stream", ex);
                    newInner = null;
                }
            }
            else
            {
                // the inner stream is a ZipFileInputStream, i.e. the data wasn't compressed
                newInner = null;
            }

            return new ThresholdInputStream(zipIS, newInner);
        }

        public class ThresholdInputStream : Stream
        {
            long counter = 0;
            ThresholdInputStream cis;
            Stream input;

            public override bool CanRead => throw new NotImplementedException();

            public override bool CanSeek => throw new NotImplementedException();

            public override bool CanWrite => throw new NotImplementedException();

            public override long Length => throw new NotImplementedException();

            public override long Position { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

            public ThresholdInputStream(Stream is1, ThresholdInputStream cis)

            {
                this.input = is1;
                this.cis = cis;
            }

            public int Read()
            {

                int b = this.input.ReadByte();
                if(b > -1)
                    Advance(1);
                return b;
            }

            public override int Read(byte[] b, int off, int len)
            {

                int cnt = input.Read(b, off, len);
                if(cnt > -1)
                    Advance(cnt);
                return cnt;

            }

            public long Skip(long n)
            {

                counter = 0;
                return input.Seek(n, SeekOrigin.Current);
            }

            public void Reset()
            {
                counter = 0;
                input.Seek(0, SeekOrigin.Begin);
            }

            public void Advance(int advance)
            {

                counter += advance;
                // check the file size first, in case we are working on uncompressed streams
                if(counter < ZipSecureFile.MAX_ENTRY_SIZE)
                {
                    if(cis == null)
                        return;
                    double ratio = (double)cis.counter / (double)counter;
                    if(ratio >= ZipSecureFile.MIN_INFLATE_RATIO)
                        return;
                }
                throw new IOException("Zip bomb detected! The file would exceed certain limits which usually indicate that the file is used to inflate memory usage and thus could pose a security risk. "
                        + "You can adjust these limits via setMinInflateRatio() and setMaxEntrySize() if you need to work with files which exceed these limits. "
                        + "Counter: " + counter + ", cis.counter: " + (cis == null ? 0 : cis.counter) + ", ratio: " + (cis == null ? 0 : ((double) cis.counter) / counter)
                        + "Limits: MIN_INFLATE_RATIO: " + ZipSecureFile.MIN_INFLATE_RATIO + ", MAX_ENTRY_SIZE: " + ZipSecureFile.MAX_ENTRY_SIZE);
            }

            public ZipEntry GetNextEntry()
            {

                if(!(input is ZipInputStream))
                {
                    throw new NotSupportedException("underlying stream is not a ZipInputStream");
                }
                counter = 0;
                return ((ZipInputStream) input).GetNextEntry();
            }

            public void CloseEntry()
            {

                if(!(input is ZipInputStream))
                {
                    throw new NotSupportedException("underlying stream is not a ZipInputStream");
                }
                counter = 0;
                ((ZipInputStream) input).CloseEntry();
            }

            public void Unread(int b)
            {

                if(!(input is PushbackInputStream))
                {
                    throw new NotSupportedException("underlying stream is not a PushbackInputStream");
                }
                if(--counter < 0)
                    counter = 0;
                ((PushbackInputStream) input).Unread(b);
            }

            public void Unread(byte[] b, int off, int len)
            {

                if(!(input is PushbackInputStream))
                {
                    throw new NotSupportedException("underlying stream is not a PushbackInputStream");
                }
                counter -= len;
                if(--counter < 0)
                    counter = 0;
                ((PushbackInputStream) input).Unread(b, off, len);
            }

            public int Available()
            {
                return (int) (input.Length - input.Position);
                //return input.Available();
            }

            public bool MarkSupported()
            {
                //return input.MarkSupported();
                return true;
            }

            public void Mark(int readlimit)
            {
                //input.Mark(readlimit);
            }

            public override void Flush()
            {
                throw new NotImplementedException();
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
        }

    }
}
