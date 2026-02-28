using System.Net;
using BenchmarkDotNet.Attributes;
using NPOI.OpenXml4Net.OPC;
using NPOI.XSSF.EventUserModel;
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
        _filePath = Path.Combine("data","test-performance.xlsx");

        // a 17MB Excel file is large so download it only when needed
        /*if (!File.Exists(_filePath))
        {
            Console.WriteLine("Downloading file...");
            new WebClient().DownloadFile(
                "https://github.com/GrapeCity/GcExcel-Java/raw/master/benchmark/files/test-performance.xlsx",
                _filePath);
            Console.WriteLine("File downloaded");
        }*/

        var copyPath = Path.Combine(Path.GetTempPath(), "test-performance-copy.xlsx");
        if (!File.Exists(copyPath))
        {
            File.Copy(_filePath, copyPath);
        }

        _loadedWorkBook = new XSSFWorkbook(copyPath);
        _memoryStream = new MemoryStream();
    }

    [Benchmark]
    public void XSSFWorkbookLoad()
    {
        var workbook = new XSSFWorkbook(_filePath, true);
        workbook.Dispose();
    }

    [Benchmark]
    public void XSSFReaderLoad()
    {
        using var pkg = OPCPackage.Open(_filePath, PackageAccess.READ);
        var reader = new XSSFReader(pkg);

        // Read shared strings table
        var sst = reader.SharedStringsTable;

        // Read styles table
        var styles = reader.StylesTable;

        // get streams of sheets data 
        var sheets = reader.GetSheetsData();
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