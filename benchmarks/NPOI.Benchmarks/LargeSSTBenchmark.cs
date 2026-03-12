using BenchmarkDotNet.Attributes;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Benchmarks;

[ShortRunJob]
[MemoryDiagnoser]
public class LargeSSTBenchmark
{
    // 36 MB .xlsx sourced from https://github.com/mini-software/MiniExcel/tree/master/benchmarks/MiniExcel.Benchmarks
    // 1,000,000 rows × 10 cols, all cells are shared strings → uniqueCount=1,000,000
    // Uncompressed sharedStrings.xml is ~31 MB, making SST the dominant parse cost.
    private string _largeFileWithSstPath;

    [GlobalSetup]
    public void GlobalSetup()
    {
        _largeFileWithSstPath = Path.Combine("data", "Test1000000x10_SharingStrings.xlsx");
    }

    /// <summary>
    /// Opens a 36 MB workbook whose sharedStrings.xml decompresses to ~31 MB
    /// (1,000,000 unique strings) and immediately disposes without reading any cells.
    /// With lazy SST loading the shared strings table is never parsed, so this
    /// benchmark represents the minimum overhead of opening the workbook.
    /// Compare with <see cref="XSSFWorkbookLargeSstLoadStrings"/> which forces
    /// SST parsing to be able to read cell values.
    /// </summary>
    [Benchmark]
    public void XSSFWorkbookLargeSstOpenDispose()
    {
        using var workbook = new XSSFWorkbook(_largeFileWithSstPath, true);
    }

    /// <summary>
    /// Opens the same 36 MB workbook and explicitly forces the SST to load by
    /// accessing <see cref="NPOI.XSSF.Model.SharedStringsTable.Count"/>.
    /// This is the baseline that shows how expensive eager DOM-based SST parsing
    /// would be; with lazy loading + streaming parser the cost is deferred and
    /// reduced in allocation compared to the old DOM path.
    /// </summary>
    [Benchmark]
    public void XSSFWorkbookLargeSstLoadStrings()
    {
        using var workbook = new XSSFWorkbook(_largeFileWithSstPath, true);
        // Force SST parse
        _ = workbook.GetSharedStringSource().Count;
    }
}
