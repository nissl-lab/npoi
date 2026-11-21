using BenchmarkDotNet.Attributes;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace NPOI.Benchmarks;

[ShortRunJob]
[MemoryDiagnoser]
public class MergedRegionsBenchmark
{
    private XSSFWorkbook _workbook;
    private XSSFSheet _sheet;

    [Params(100, 1000, 5000)]
    public int RegionCount { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _workbook = new XSSFWorkbook();
        _sheet = (XSSFSheet)_workbook.CreateSheet("test");
        // Add non-overlapping merged regions (each row merges columns 0-1)
        for (int i = 0; i < RegionCount; i++)
        {
            _sheet.AddMergedRegionUnsafe(new CellRangeAddress(i, i, 0, 1));
        }
    }

    [Benchmark]
    public int ReadMergedRegions()
    {
        // Simulates repeated reads (e.g., during row copy, auto-size, validation)
        int count = 0;
        for (int i = 0; i < 10; i++)
        {
            count = _sheet.MergedRegions.Count;
        }
        return count;
    }

    [Benchmark]
    public int AddMergedRegions()
    {
        // Simulates building a sheet with many merged regions (validated)
        var tempSheet = (XSSFSheet)_workbook.CreateSheet();
        for (int i = 0; i < RegionCount; i++)
        {
            tempSheet.AddMergedRegion(new CellRangeAddress(i, i, 0, 1));
        }
        int count = tempSheet.MergedRegions.Count;
        _workbook.RemoveSheetAt(_workbook.GetSheetIndex(tempSheet));
        return count;
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        _workbook?.Dispose();
    }
}
