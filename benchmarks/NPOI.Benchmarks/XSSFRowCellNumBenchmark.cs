using BenchmarkDotNet.Attributes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOI.Benchmarks;

[ShortRunJob]
[MemoryDiagnoser]
public class XSSFRowCellNumBenchmark
{
    private XSSFWorkbook _workbook;
    private IRow _sparseRow;
    private IRow _denseRow;

    [GlobalSetup]
    public void Setup()
    {
        _workbook = new XSSFWorkbook();
        var sheet = _workbook.CreateSheet("test");

        // Sparse row: cells at widely separated columns
        _sparseRow = sheet.CreateRow(0);
        _sparseRow.CreateCell(0).SetCellValue("first");
        _sparseRow.CreateCell(500).SetCellValue("mid");
        _sparseRow.CreateCell(1000).SetCellValue("last");

        // Dense row: 200 contiguous cells (typical spreadsheet)
        _denseRow = sheet.CreateRow(1);
        for (int i = 0; i < 200; i++)
        {
            _denseRow.CreateCell(i).SetCellValue(i);
        }
    }

    [Benchmark]
    public int SparseRow_FirstLastCellNum_10000x()
    {
        // Simulates tight loop pattern: for (j = row.FirstCellNum; j <= row.LastCellNum; j++)
        int sum = 0;
        for (int i = 0; i < 10_000; i++)
        {
            sum += _sparseRow.FirstCellNum + _sparseRow.LastCellNum;
        }
        return sum;
    }

    [Benchmark]
    public int DenseRow_FirstLastCellNum_10000x()
    {
        int sum = 0;
        for (int i = 0; i < 10_000; i++)
        {
            sum += _denseRow.FirstCellNum + _denseRow.LastCellNum;
        }
        return sum;
    }

    [Benchmark]
    public int DenseRow_IterateCellRange()
    {
        // Pattern seen in copy/shift operations: iterate FirstCellNum..LastCellNum
        int count = 0;
        for (int j = _denseRow.FirstCellNum; j < _denseRow.LastCellNum; j++)
        {
            ICell cell = _denseRow.GetCell(j);
            if (cell != null) count++;
        }
        return count;
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        _workbook?.Dispose();
    }
}
