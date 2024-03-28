using System;
using System.Collections;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;

namespace NPOI.OpenXml4Net.Util
{
    /// <summary>
    /// An Interface to make getting the different bits
    ///  of a Zip File easy.
    /// Allows you to get at the ZipEntries, without
    ///  needing to worry about ZipFile vs ZipInputStream
    ///  being annoyingly very different.
    /// </summary>
    public interface ZipEntrySource
    {
        /// <summary>
        /// Returns an Enumeration of all the Entries
        /// </summary>
        IEnumerator Entries { get; }

        /// <summary>
        /// Returns an InputStream of the decompressed
        ///  data that makes up the entry
        /// </summary>
        Stream GetInputStream(ZipEntry entry);

        /// <summary>
        /// Indicates we are done with reading, and
        ///  resources may be freed
        /// </summary>
        void Close();

        /// <summary>
    /// Has close been called already?
    /// </summary>
        bool IsClosed { get; }
    }
}
