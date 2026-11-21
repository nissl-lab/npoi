using System.IO.Compression;

namespace NPOI.Benchmarks;

/// <summary>
/// Utility helpers for inspecting zip/xlsx entries without fully extracting the archive.
/// </summary>
internal static class ZipUtils
{
    /// <summary>
    /// Returns the uncompressed byte length of the <c>xl/styles.xml</c> entry
    /// inside an xlsx file, or -1 if the entry is not found.
    /// </summary>
    public static long GetStylesXmlUncompressedSize(string xlsxPath)
    {
        using var zip = ZipFile.OpenRead(xlsxPath);
        var entry = zip.GetEntry("xl/styles.xml");
        return entry?.Length ?? -1L;
    }
}
