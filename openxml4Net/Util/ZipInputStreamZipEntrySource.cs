using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Collections;
using ICSharpCode.SharpZipLib.Zip;

namespace NPOI.OpenXml4Net.Util
{
    /**
     * Provides a way to get at all the ZipEntries
     *  from a ZipInputStream, as many times as required.
     * Allows a ZipInputStream to be treated much like
     *  a ZipFile, for a price in terms of memory.
     * Be sure to call {@link #close()} as soon as you're
     *  done, to free up that memory!
     */
    public class ZipInputStreamZipEntrySource : ZipEntrySource
    {
        private List<FakeZipEntry> zipEntries;

        /**
         * Reads all the entries from the ZipInputStream
         *  into memory, and closes the source stream.
         * We'll then eat lots of memory, but be able to
         *  work with the entries at-will.
         */
        public ZipInputStreamZipEntrySource(ZipInputStream inp)
            : this(inp, null)
        {
        }

        /**
         * As {@link #ZipInputStreamZipEntrySource(ZipInputStream)}, but additionally
         * passes a counter over the raw (compressed) source stream so that each entry
         * can be checked for a decompression-bomb ratio while it is buffered. This is
         * required because streamed / data-descriptor entries report Size == -1, so the
         * uncompressed size is not known up-front and the ratio must be derived from the
         * compressed bytes actually consumed.
         */
        public ZipInputStreamZipEntrySource(ZipInputStream inp, CountingStream compressedCounter)
        {
            zipEntries = new List<FakeZipEntry>();

            bool going = true;
            //if(inp.Position != 0)
            //    inp.Position = 0;
            while (going)
            {
                ZipEntry zipEntry = inp.GetNextEntry();
                if (zipEntry == null)
                {
                    going = false;
                }
                else
                {
                    FakeZipEntry entry = new FakeZipEntry(zipEntry, inp, compressedCounter);
                    //inp.Close();

                    zipEntries.Add(entry);
                }
            }
            inp.Close();
        }

        public IEnumerator Entries
        {
            get
            {
                return new EntryEnumerator(zipEntries);
            }
        }

        public Stream GetInputStream(ZipEntry zipEntry)
        {
            FakeZipEntry entry = (FakeZipEntry)zipEntry;
            return entry.GetInputStream();
        }

        public void Close()
        {
            // Free the memory
            zipEntries = null;
        }

        public bool IsClosed
        {
            get { return zipEntries == null; }
        }
        /**
         * Why oh why oh why are Iterator and Enumeration
         *  still not compatible?
         */
        internal sealed class EntryEnumerator : IEnumerator
        {
            private List<FakeZipEntry>.Enumerator iterator;

            internal EntryEnumerator(List<FakeZipEntry> zipEntries)
            {
                iterator = zipEntries.GetEnumerator();
            }

            public bool MoveNext()
            {
                return iterator.MoveNext();
            }

            public object Current
            {
                get
                {
                    return iterator.Current;
                }
            }

            #region IEnumerator Members


            public void Reset()
            {
                throw new NotImplementedException();
            }

            #endregion
        }

        /**
         * So we can close the real zip entry and still
         *  effectively work with it.
         * Holds the (decompressed!) data in memory, so
         *  close this as soon as you can! 
         */
        public class FakeZipEntry : ZipEntry
        {
            private byte[] data;

            public FakeZipEntry(ZipEntry entry, ZipInputStream inp)
                : this(entry, inp, null)
            {
            }

            public FakeZipEntry(ZipEntry entry, ZipInputStream inp, CountingStream compressedCounter) : base(entry.Name)
            {

                // Grab the de-compressed contents for later
                MemoryStream baos;

                long entrySize = entry.Size;

                if (entrySize != -1)
                {
                    if (entrySize >= Int32.MaxValue)
                    {
                        throw new IOException("ZIP entry size is too large");
                    }

                    baos = new MemoryStream((int)entrySize);
                }
                else
                {
                    baos = new MemoryStream();
                }

                // Baseline of compressed bytes consumed before this entry's data is read,
                // so we can measure this entry's compressed size even when the header omits
                // it (streamed / data-descriptor entries with Size == -1).
                long compressedStart = compressedCounter != null ? compressedCounter.BytesRead : -1;

                byte[] buffer = new byte[4096];
                long decompressed = 0;
                int read = 0;
                while ((read = inp.Read(buffer, 0, buffer.Length)) > 0)
                {
                    decompressed += read;

                    // Prefer the header's compressed size when present and trustworthy;
                    // otherwise fall back to the raw bytes actually consumed for this entry.
                    long compressed = compressedStart >= 0
                        ? compressedCounter.BytesRead - compressedStart
                        : entry.CompressedSize;

                    ZipSecureFile.CheckThreshold(decompressed, compressed);

                    baos.Write(buffer, 0, read);
                }

                data = baos.ToArray();
            }

            public Stream GetInputStream()
            {
                return new MemoryStream(data);
            }
        }
    }

}
