using BenchmarkDotNet.Attributes;
using NPOI.XSSF.UserModel;
using System.IO;

namespace NPOI.Benchmarks;

/// <summary>
/// Measures the write-side cost of <see cref="NPOI.XSSF.Model.SharedStringsTable"/> serialization.
/// Creates a workbook with <see cref="WriteRowCount"/> rows of unique string cells
/// (forcing SST dirty) and writes it to a <see cref="MemoryStream"/>.
/// </summary>
[ShortRunJob]
[MemoryDiagnoser]
public class SSTWriteBenchmark
{
    [Params(10_000, 100_000)]
    public int WriteRowCount { get; set; }

    /// <summary>
    /// Creates a workbook with <see cref="WriteRowCount"/> rows of unique string cells
    /// (forcing SST dirty) and writes it to a MemoryStream.
    /// Measures the write-side cost of SharedStringsTable serialization.
    /// </summary>
    [Benchmark]
    public void XSSFWorkbookWriteLargeSst()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet();
        for (int i = 0; i < WriteRowCount; i++)
            sheet.CreateRow(i).CreateCell(0).SetCellValue("UniqueString_" + i);
        using var ms = new MemoryStream();
        workbook.Write(ms, leaveOpen: true);
    }
}
