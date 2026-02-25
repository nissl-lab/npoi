using BenchmarkDotNet.Attributes;
using NPOI.XSSF.UserModel;

namespace NPOI.Benchmarks;

[MemoryDiagnoser]
public class AddRowBenchmark
{
    [Params(1_000, 10_000)]
    public int RowCount { get; set; }

    [Benchmark]
    public void AddRowsSequentially()
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Sheet1");

        for (var i = 0; i < RowCount; i++)
        {
            sheet.CreateRow(i);
        }

        workbook.Close();
    }

    [Benchmark]
    public void AddRowsInReverse()
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Sheet1");

        for (var i = RowCount - 1; i >= 0; i--)
        {
            sheet.CreateRow(i);
        }

        workbook.Close();
    }
}
