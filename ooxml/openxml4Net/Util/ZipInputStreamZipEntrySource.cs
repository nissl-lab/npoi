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
public class ZipInputStreamZipEntrySource:ZipEntrySource {
    private List<FakeZipEntry> zipEntries;
	
	/**
	 * Reads all the entries from the ZipInputStream 
	 *  into memory, and closes the source stream.
	 * We'll then eat lots of memory, but be able to
	 *  work with the entries at-will.
	 */
	public ZipInputStreamZipEntrySource(ZipInputStream inp){
        zipEntries = new List<FakeZipEntry>();
		
		bool going = true;
		while(going) {
			ZipEntry zipEntry = inp.GetNextEntry();
			if(zipEntry == null) {
				going = false;
			} else {
				FakeZipEntry entry = new FakeZipEntry(zipEntry, inp);
				//inp.Close();

                zipEntries.Add(entry);
			}
		}
		inp.Close();
	}

    public IEnumerator Entries
    {
        get{
		return new EntryEnumerator(zipEntries);
        }
	}
	
	public Stream GetInputStream(ZipEntry zipEntry) {
        FakeZipEntry entry = (FakeZipEntry)zipEntry;
		return entry.GetInputStream();
	}
	
	public void Close() {
		// Free the memory
		zipEntries = null;
	}
	
	/**
	 * Why oh why oh why are Iterator and Enumeration
	 *  still not compatible?
	 */
	internal class EntryEnumerator:IEnumerator {
        private List<FakeZipEntry>.Enumerator iterator;

        internal EntryEnumerator(List<FakeZipEntry> zipEntries)
        {
			iterator = zipEntries.GetEnumerator();
		}
		
		public bool MoveNext() {
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
	public class FakeZipEntry : ZipEntry {
		private byte[] data;
		
		public FakeZipEntry(ZipEntry entry, ZipInputStream inp):base(entry.Name)
        {

            // Grab the de-compressed contents for later
            MemoryStream baos = new MemoryStream();

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

			byte[] buffer = new byte[4096];
			int read = 0;
			while( (read = inp.Read(buffer,0,buffer.Length)) > 0 ) {
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
