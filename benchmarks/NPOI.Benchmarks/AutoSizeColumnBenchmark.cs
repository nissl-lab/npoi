using BenchmarkDotNet.Attributes;
using NPOI.SS.UserModel;
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

    [Params(1_000)]
    public int RowCount { get; set; }

    [Params(5)]
    public int ColumnCount { get; set; }

    [IterationSetup]
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

    [Benchmark]
    public void AutoSizeColumn()
    {
        for (var col = 1; col <= ColumnCount; col++)
        {
            sheet1.AutoSizeColumn(col);
        }
    }

    [IterationCleanup]
    public void Cleanup()
    {
        workbook.Dispose();
    }
}