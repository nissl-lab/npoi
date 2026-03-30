using BenchmarkDotNet.Attributes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOI.Benchmarks;

[ShortRunJob]
[MemoryDiagnoser]
public class CellValueCachingBenchmark
{
    private XSSFWorkbook _workbook;
    private XSSFSheet _sheet;
    private ICell _numericCell;
    private ICell _stringCell;
    
    [GlobalSetup]
    public void GlobalSetup()
    {
        _workbook = new XSSFWorkbook();
        _sheet = (XSSFSheet)_workbook.CreateSheet("Benchmark");
        
        // Create a numeric cell
        var row1 = _sheet.CreateRow(0);
        _numericCell = row1.CreateCell(0);
        _numericCell.SetCellValue(123.456);
        
        // Create a string cell
        var row2 = _sheet.CreateRow(1);
        _stringCell = row2.CreateCell(0);
        _stringCell.SetCellValue("Hello World Test String");
    }
    
    [GlobalCleanup]
    public void GlobalCleanup()
    {
        _workbook?.Dispose();
    }
    
    [Benchmark]
    public double RepeatedNumericAccess()
    {
        double sum = 0;
        // Access the same cell's numeric value 1000 times
        for (int i = 0; i < 1000; i++)
        {
            sum += _numericCell.NumericCellValue;
        }
        return sum;
    }
    
    [Benchmark]
    public string RepeatedStringAccess()
    {
        var sb = new System.Text.StringBuilder();
        // Access the same cell's string value 1000 times
        for (int i = 0; i < 1000; i++)
        {
            sb.Append(_stringCell.StringCellValue);
        }
        return sb.ToString();
    }
    
    [Benchmark]
    public double RepeatedRichStringAccess()
    {
        int totalLength = 0;
        // Access the same cell's rich string value 1000 times
        for (int i = 0; i < 1000; i++)
        {
            totalLength += _stringCell.RichStringCellValue.String.Length;
        }
        return totalLength;
    }
}