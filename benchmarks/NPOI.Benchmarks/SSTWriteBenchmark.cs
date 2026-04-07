using BenchmarkDotNet.Attributes;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;
using System.IO;

namespace NPOI.Benchmarks;

/// <summary>
/// Measures the write-side cost of <see cref="NPOI.XSSF.Model.SharedStringsTable"/> serialization.
/// Creates a workbook with <see cref="WriteRowCount"/> rows of unique string cells
/// (forcing SST dirty) and writes it to a <see cref="MemoryStream"/>.
/// Compares the new direct-streaming path (<see cref="SharedStringsTable.UseDirectWrite"/> = true)
/// against the legacy <c>_sstDoc.Save()</c> path.
/// </summary>
[ShortRunJob]
[MemoryDiagnoser]
public class SSTWriteBenchmark
{
    [Params(10_000, 100_000)]
    public int WriteRowCount { get; set; }

    /// <summary>
    /// New path: streams XML directly from the <c>strings</c> list, bypassing
    /// <c>_sstDoc.Save()</c> and the intermediate <c>CT_Sst.si</c> list.
    /// </summary>
    [Benchmark]
    public void XSSFWorkbookWriteLargeSstDirectWrite()
    {
        SharedStringsTable.UseDirectWrite = true;
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet();
        for (int i = 0; i < WriteRowCount; i++)
            sheet.CreateRow(i).CreateCell(0).SetCellValue("UniqueString_" + i);
        using var ms = new MemoryStream();
        workbook.Write(ms, leaveOpen: true);
    }

    /// <summary>
    /// Legacy path: rebuilds <c>sst.si</c> from the <c>strings</c> list and delegates
    /// to <c>_sstDoc.Save()</c>, matching the behavior before the direct-write optimization.
    /// </summary>
    [Benchmark]
    public void XSSFWorkbookWriteLargeSstLegacy()
    {
        SharedStringsTable.UseDirectWrite = false;
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet();
        for (int i = 0; i < WriteRowCount; i++)
            sheet.CreateRow(i).CreateCell(0).SetCellValue("UniqueString_" + i);
        using var ms = new MemoryStream();
        workbook.Write(ms, leaveOpen: true);
    }
}
