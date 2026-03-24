using BenchmarkDotNet.Attributes;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace NPOI.Benchmarks;

[MemoryDiagnoser]
public class AutoSizeColumnBenchmark
{
    private static readonly string[] lorem = @"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut
labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip
ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat
nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id
est laborum.".Split(' ', '\r', '\n');

    private XSSFWorkbook workbook;
    private ISheet sheet1;

    [Params(1_000, 10_000)]
    public int RowCount { get; set; }

    [Params(5)]
    public int ColumnCount { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        var ipsum = 0;

        workbook = new XSSFWorkbook();
        sheet1 = workbook.CreateSheet("Sheet1");

        for (var rowNum = 1; rowNum <= RowCount; rowNum++)
        {
            var row = sheet1.CreateRow(rowNum);
            for (int col = 1; col <= ColumnCount; col++)
            {
                row.CreateCell(col).SetCellValue(lorem[ipsum++ % lorem.Length]);
            }
        }
    }

    [Benchmark(Baseline = true)]
    public void AutoSizeColumn()
    {
        for (var col = 1; col <= ColumnCount; col++)
        {
            sheet1.AutoSizeColumn(col);
        }
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        workbook.Dispose();
    }
}

[MemoryDiagnoser]
public class AutoSizeColumnWithMergedRegionsBenchmark
{
    private static readonly string[] lorem = @"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut
labore et dolore magna aliqua.".Split(' ', '\r', '\n');

    private XSSFWorkbook workbook;
    private ISheet sheet1;

    [Params(1_000)]
    public int RowCount { get; set; }

    [Params(5)]
    public int ColumnCount { get; set; }

    [Params(0, 100)]
    public int MergedRegionCount { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        var ipsum = 0;

        workbook = new XSSFWorkbook();
        sheet1 = workbook.CreateSheet("Sheet1");

        for (var rowNum = 0; rowNum < RowCount; rowNum++)
        {
            var row = sheet1.CreateRow(rowNum);
            for (int col = 0; col < ColumnCount; col++)
            {
                row.CreateCell(col).SetCellValue(lorem[ipsum++ % lorem.Length]);
            }
        }

        // Add merged regions spanning 2 columns on evenly-spaced rows
        if (MergedRegionCount > 0)
        {
            int step = Math.Max(1, RowCount / MergedRegionCount);
            int added = 0;
            for (int r = 0; r < RowCount && added < MergedRegionCount; r += step, added++)
            {
                sheet1.AddMergedRegion(new CellRangeAddress(r, r, 0, 1));
            }
        }
    }

    [Benchmark]
    public void AutoSizeColumnWithMergedCells()
    {
        for (var col = 0; col < ColumnCount; col++)
        {
            sheet1.AutoSizeColumn(col, true);
        }
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        workbook.Dispose();
    }
}