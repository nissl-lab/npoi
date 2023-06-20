using BenchmarkDotNet.Attributes;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace NPOI.Benchmarks;

[MemoryDiagnoser]
public class RangeValuesBenchmark
{
    private static readonly MemoryStream _memoryStream = new();

    [Params(1_000)]
    public int RowCount { get; set; }

    [Params(30)]
    public int ColumnCount { get; set; }

    [Benchmark]
    public void Double()
    {
        var workbook = new XSSFWorkbook();
        var worksheet = workbook.CreateSheet("poi");

        for (var r = 0; r < RowCount; r++)
        {
            var row = worksheet.CreateRow(r);
            for (var c = 0; c < ColumnCount; c++)
            {
                row.CreateCell(c).SetCellValue(r + c);
            }
        }

        for (var r = 0; r < RowCount; r++)
        {
            var row = worksheet.GetRow(r);
            for (var c = 0; c < ColumnCount; c++)
            {
                var d = row.GetCell(c).NumericCellValue;
            }
        }

        WriteFileAndClose(workbook);
    }

    [Benchmark]
    public void String()
    {
        var workbook = new XSSFWorkbook();
        var worksheet = workbook.CreateSheet("poi");

        var random = new Random();
        const string alphaNumericString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        for (var r = 0; r < RowCount; r++)
        {
            var row = worksheet.CreateRow(r);
            for (var c = 0; c < ColumnCount; c++)
            {
                row.CreateCell(c).SetCellValue(alphaNumericString[random.Next(25)].ToString());
            }
        }

        for (var r = 0; r < RowCount; r++)
        {
            var row = worksheet.GetRow(r);
            for (var c = 0; c < ColumnCount; c++)
            {
                var s = row.GetCell(c).StringCellValue;
            }
        }

        WriteFileAndClose(workbook);
    }

    [Benchmark]
    public void Date()
    {
        var workbook = new XSSFWorkbook();
        var worksheet = workbook.CreateSheet("poi");

        for (var r = 0; r < RowCount; r++)
        {
            var row = worksheet.CreateRow(r);
            for (var c = 0; c < ColumnCount; c++)
            {
                row.CreateCell(c).SetCellValue(DateTime.Now);
            }
        }

        for (var r = 0; r < RowCount; r++)
        {
            var row = worksheet.GetRow(r);
            for (var c = 0; c < ColumnCount; c++)
            {
                var d = row.GetCell(c).DateCellValue;
            }
        }

        WriteFileAndClose(workbook);
    }

    [Benchmark]
    public void Formulas()
    {
        var workbook = new XSSFWorkbook();
        var worksheet = workbook.CreateSheet("poi");

        for (var r = 0; r < RowCount; r++)
        {
            var row = worksheet.CreateRow(r);
            for (var c = 0; c < 2; c++)
            {
                row.CreateCell(c).SetCellValue(r + c);
            }
        }

        for (var r = 0; r < RowCount; r++)
        {
            var row = worksheet.GetRow(r);
            for (var c = 2; c < ColumnCount + 2; c++)
            {
                var cell = row.CreateCell(c);
                var reference1 = new CellReference(r, c - 2);
                var reference2 = new CellReference(r, c - 1);
                cell.CellFormula = $"SUM({reference1.FormatAsString()}, {reference2.FormatAsString()})";
            }
        }

        workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

        WriteFileAndClose(workbook);
    }

    private static void WriteFileAndClose(XSSFWorkbook workbook)
    {
        _memoryStream.Seek(0, SeekOrigin.Begin);
        workbook.Write(_memoryStream, leaveOpen: true);
        workbook.Close();
    }
}