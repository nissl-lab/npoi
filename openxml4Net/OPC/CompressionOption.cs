using System;
using System.Collections.Generic;
using System.Text;
using ICSharpCode.SharpZipLib.Zip.Compression;

namespace NPOI.OpenXml4Net.OPC
{
    public enum CompressionOption : int
    {
        Fast = Deflater.BEST_SPEED,
        Maximum = Deflater.BEST_COMPRESSION,
        Normal = Deflater.DEFAULT_COMPRESSION,
        NotCompressed = Deflater.NO_COMPRESSION
    }
}
