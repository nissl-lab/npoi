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

    // 36 MB .xlsx sourced from https://github.com/mini-software/MiniExcel/tree/master/benchmarks/MiniExcel.Benchmarks
    // 1,000,000 rows × 10 cols, all cells are shared strings → uniqueCount=1,000,000
    // Uncompressed sharedStrings.xml is ~31 MB, making SST the dominant parse cost.
    private string _largeFileWithSstPath;

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

        _largeFileWithSstPath = Path.Combine("data", "Test1000000x10_SharingStrings.xlsx");
    }

    [Benchmark]
    public void XSSFWorkbookLoad()
    {
        var workbook = new XSSFWorkbook(_filePath, true);
        workbook.Dispose();
    }

    /// <summary>
    /// Opens a 36 MB workbook whose sharedStrings.xml decompresses to ~31 MB
    /// (1,000,000 unique strings) and immediately disposes without reading any cells.
    /// With lazy SST loading the shared strings table is never parsed, so this
    /// benchmark represents the minimum overhead of opening the workbook.
    /// Compare with <see cref="XSSFWorkbookLargeSstLoadStrings"/> which forces
    /// SST parsing to be able to read cell values.
    /// </summary>
    [Benchmark]
    public void XSSFWorkbookLargeSstOpenDispose()
    {
        using var workbook = new XSSFWorkbook(_largeFileWithSstPath, true);
    }

    /// <summary>
    /// Opens the same 36 MB workbook and explicitly forces the SST to load by
    /// accessing <see cref="NPOI.XSSF.Model.SharedStringsTable.Count"/>.
    /// This is the baseline that shows how expensive eager DOM-based SST parsing
    /// would be; with lazy loading + streaming parser the cost is deferred and
    /// reduced in allocation compared to the old DOM path.
    /// </summary>
    [Benchmark]
    public void XSSFWorkbookLargeSstLoadStrings()
    {
        using var workbook = new XSSFWorkbook(_largeFileWithSstPath, true);
        // Force SST parse
        _ = workbook.GetSharedStringSource().Count;
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