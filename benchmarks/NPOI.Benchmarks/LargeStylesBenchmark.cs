using BenchmarkDotNet.Attributes;
using NPOI.XSSF.UserModel;

namespace NPOI.Benchmarks;

/// <summary>
/// Benchmarks for workbooks that contain a large <c>xl/styles.xml</c>.
/// The fixture file is generated automatically on first run and reused on
/// subsequent runs as long as its uncompressed styles.xml is &gt;= 20 MB.
/// </summary>
[ShortRunJob]
[MemoryDiagnoser]
public class LargeStylesBenchmark
{
    private string _largeFileWithStylesPath = string.Empty;

    [GlobalSetup]
    public void GlobalSetup()
    {
        _largeFileWithStylesPath = Path.Combine("data", "HugeStyles.xlsx");
    }

    /// <summary>
    /// Opens and immediately disposes the workbook without touching any style API.
    /// With lazy <c>styles.xml</c> loading this should avoid parsing the style table.
    /// </summary>
    [Benchmark]
    public void XSSFWorkbookLargeStylesOpenDispose()
    {
        using var workbook = new XSSFWorkbook(_largeFileWithStylesPath, readOnly: true);
    }

    /// <summary>
    /// Opens the workbook and forces the style table to load by calling
    /// <see cref="XSSFWorkbook.CreateCellStyle"/> and
    /// <see cref="XSSFWorkbook.CreateDataFormat"/>.
    /// </summary>
    [Benchmark]
    public void XSSFWorkbookLargeStylesForceLoad()
    {
        using var workbook = new XSSFWorkbook(_largeFileWithStylesPath, readOnly: true);
        _ = workbook.CreateCellStyle();
        _ = workbook.CreateDataFormat();
    }

    /// <summary>
    /// Opens the workbook in read-write mode and writes it to a <see cref="MemoryStream"/>.
    /// The workbook is opened in write mode (not read-only) so that the full save
    /// path is exercised, including any styles serialisation.  Once the
    /// "copy-through styles.xml when not dirty" optimisation is in place this
    /// should avoid re-serialising <c>xl/styles.xml</c> altogether.
    /// </summary>
    [Benchmark]
    public void XSSFWorkbookLargeStylesOpenWrite()
    {
        // readOnly: false is intentional – we want to exercise the full write path.
        using var workbook = new XSSFWorkbook(_largeFileWithStylesPath, readOnly: false);
        using var ms = new MemoryStream();
        workbook.Write(ms, leaveOpen: true);
    }
}
