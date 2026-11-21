using BenchmarkDotNet.Attributes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOI.Benchmarks;

[ShortRunJob]
[MemoryDiagnoser]
[RankColumn]
public class RowCellAccessBenchmark
{
    private XSSFWorkbook _workbook;
    private XSSFSheet _sheet;
    private const int RowCount = 1000;
    private const int ColumnCount = 50;

    [GlobalSetup]
    public void GlobalSetup()
    {
        _workbook = new XSSFWorkbook();
        _sheet = (XSSFSheet)_workbook.CreateSheet("Benchmark");

        for (int row = 0; row < RowCount; row++)
        {
            IRow excelRow = _sheet.CreateRow(row);
            for (int col = 0; col < ColumnCount; col++)
            {
                excelRow.CreateCell(col).SetCellValue($"Data_{row}_{col}");
            }
        }
    }

    [GlobalCleanup]
    public void GlobalCleanup()
    {
        _workbook?.Dispose();
    }

    private static int _sheetCounter = 0;

    [Benchmark]
    public void CreateRow()
    {
        var sheet = (XSSFSheet)_workbook.CreateSheet($"S{System.Threading.Interlocked.Increment(ref _sheetCounter):D4}");
        for (int row = 0; row < 100; row++)
        {
            IRow excelRow = sheet.CreateRow(row);
            for (int col = 0; col < 20; col++)
            {
                excelRow.CreateCell(col).SetCellValue($"Data_{row}_{col}");
            }
        }
    }

    [Benchmark]
    public void CreateCell_AppendOrder()
    {
        var row = _sheet.CreateRow(RowCount + 1);
        for (int col = 0; col < ColumnCount; col++)
        {
            row.CreateCell(col).SetCellValue($"New_{col}");
        }
        _sheet.RemoveRow(row);
    }

    [Benchmark]
    public void CreateCell_RandomOrder()
    {
        var row = _sheet.CreateRow(RowCount + 2);
        for (int col = ColumnCount - 1; col >= 0; col--)
        {
            row.CreateCell(col).SetCellValue($"New_{col}");
        }
        _sheet.RemoveRow(row);
    }

    [Benchmark]
    public void GetCell_Existing()
    {
        int sum = 0;
        for (int row = 0; row < Math.Min(100, RowCount); row++)
        {
            var excelRow = _sheet.GetRow(row);
            for (int col = 0; col < Math.Min(20, ColumnCount); col++)
            {
                var cell = excelRow.GetCell(col);
                if (cell != null)
                    sum += cell.ColumnIndex;
            }
        }
    }

    [Benchmark]
    public void IterateAllRows()
    {
        int count = 0;
        foreach (IRow row in _sheet)
        {
            count += row.RowNum;
        }
    }

    [Benchmark]
    public void IterateAllCells()
    {
        int count = 0;
        foreach (IRow row in _sheet)
        {
            foreach (ICell cell in row)
            {
                count++;
            }
        }
    }

    [Benchmark]
    public void IterateAllCellsWithGetCell()
    {
        int count = 0;
        for (int row = 0; row < Math.Min(100, RowCount); row++)
        {
            var excelRow = _sheet.GetRow(row);
            if (excelRow == null) continue;
            for (int col = 0; col < Math.Min(20, ColumnCount); col++)
            {
                var cell = excelRow.GetCell(col);
                if (cell != null) count++;
            }
        }
    }

    [Benchmark]
    public void GetRow_ByIndex()
    {
        for (int row = 0; row < Math.Min(100, RowCount); row++)
        {
            var excelRow = _sheet.GetRow(row);
        }
    }

    [Benchmark]
    public void ReadCellValues()
    {
        double sum = 0;
        for (int row = 0; row < Math.Min(100, RowCount); row++)
        {
            var excelRow = _sheet.GetRow(row);
            if (excelRow == null) continue;
            foreach (ICell cell in excelRow)
            {
                if (cell.CellType == CellType.String)
                {
                    sum += cell.StringCellValue.Length;
                }
            }
        }
    }

    [Benchmark]
    public void DeleteCell()
    {
        var row = _sheet.CreateRow(RowCount + 3);
        for (int col = 0; col < ColumnCount; col++)
        {
            row.CreateCell(col).SetCellValue($"ToDelete_{col}");
        }
        for (int col = 0; col < ColumnCount; col++)
        {
            row.RemoveCell(row.GetCell(col));
        }
        _sheet.RemoveRow(row);
    }

    [Benchmark]
    public void IterateRowCells()
    {
        int count = 0;
        var row = _sheet.GetRow(0);
        if (row != null)
        {
            foreach (var cell in row)
            {
                count++;
            }
        }
    }
}
