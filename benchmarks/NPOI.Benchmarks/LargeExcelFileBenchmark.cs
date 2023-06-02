using System.Net;
using BenchmarkDotNet.Attributes;
using NPOI.XSSF.UserModel;

namespace NPOI.Benchmarks;

[ShortRunJob]
[MemoryDiagnoser]
public class LargeExcelFileBenchmark
{
    private static XSSFWorkbook _loadedWorkBook;
    private static MemoryStream _memoryStream;
    private string _filePath;

    [GlobalSetup]
    public void GlobalSetup()
    {
        // a 17MB Excel file is large so download it only when needed
        _filePath = Path.Combine(Path.GetTempPath(), "test-performance.xlsx");
        if (!File.Exists(_filePath))
        {
            Console.WriteLine("Downloading file...");
            new WebClient().DownloadFile(
                "https://github.com/GrapeCity/GcExcel-Java/raw/master/benchmark/files/test-performance.xlsx",
                _filePath);
            Console.WriteLine("File downloaded");
        }

        var copyPath = Path.Combine(Path.GetTempPath(), "test-performance-copy.xlsx");
        if (!File.Exists(copyPath))
        {
            File.Copy(_filePath, copyPath);
        }

        _loadedWorkBook = new XSSFWorkbook(copyPath);
        _memoryStream = new MemoryStream();
    }

    [Benchmark]
    public void Load()
    {
        var workbook = new XSSFWorkbook(_filePath);
        workbook.Dispose();
    }

    [Benchmark]
    public void Write()
    {
        _memoryStream.Seek(0, SeekOrigin.Begin);
        _loadedWorkBook.Write(_memoryStream, leaveOpen: true);
    }

    [Benchmark]
    public void Evaluate()
    {
        _loadedWorkBook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();
    }
}