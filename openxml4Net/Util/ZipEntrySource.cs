using System;
using System.Collections;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;

namespace NPOI.OpenXml4Net.Util
{
    /**
     * An Interface to make getting the different bits
     *  of a Zip File easy.
     * Allows you to get at the ZipEntries, without
     *  needing to worry about ZipFile vs ZipInputStream
     *  being annoyingly very different.
     */
    public interface ZipEntrySource
    {
        /**
         * Returns an Enumeration of all the Entries
         */
        IEnumerator Entries { get; }

        /**
         * Returns an InputStream of the decompressed 
         *  data that makes up the entry
         */
        Stream GetInputStream(ZipEntry entry);

        /**
         * Indicates we are done with reading, and 
         *  resources may be freed
         */
        void Close();
    }
}
