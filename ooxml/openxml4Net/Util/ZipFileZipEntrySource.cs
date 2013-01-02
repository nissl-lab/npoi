using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;

namespace NPOI.OpenXml4Net.Util
{
    /**
     * A ZipEntrySource wrapper around a ZipFile.
     * Should be as low in terms of memory as a
     *  normal ZipFile implementation is.
     */
    public class ZipFileZipEntrySource : ZipEntrySource
    {
        private ZipFile zipArchive;
        public ZipFileZipEntrySource(ZipFile zipFile)
        {
            this.zipArchive = zipFile;
        }

        public void Close()
        {
            zipArchive.Close();
        }

        public IEnumerator Entries
        {
            get
            {
                return zipArchive.GetEnumerator();
            }
        }

        public Stream GetInputStream(ZipEntry entry)
        {
            Stream s = zipArchive.GetInputStream(entry);
            return s;
        }
    }
}
