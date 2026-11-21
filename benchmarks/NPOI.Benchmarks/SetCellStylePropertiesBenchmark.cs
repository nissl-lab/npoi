using BenchmarkDotNet.Attributes;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace NPOI.Benchmarks;

/// <summary>
/// Compares the legacy O(n) linear style scan (before) against the new O(1)
/// StyleKey/StyleCache lookup (after) when calling
/// <see cref="CellUtil.SetCellStyleProperties"/> on many cells.
///
/// Run with:
///   dotnet run -c Release --project benchmarks/NPOI.Benchmarks/ -- --filter *SetCellStyleProperties*
/// </summary>
[ShortRunJob]
[MemoryDiagnoser]
public class SetCellStylePropertiesBenchmark
{
    private HSSFWorkbook _hssfWorkbook;
    private XSSFWorkbook _xssfWorkbook;
    private ICell[] _hssfCells;
    private ICell[] _xssfCells;
    private Dictionary<string, object> _singleStyleProps;
    private Dictionary<string, object>[] _multiStyleProps;

    /// <summary>
    /// Number of cells to style in each benchmark iteration.
    /// </summary>
    [Params(100, 500)]
    public int CellCount { get; set; }

    /// <summary>
    /// When <c>true</c> (the new "after" path) uses the O(1) StyleKey/StyleCache lookup.
    /// When <c>false</c> (the old "before" path) performs the O(n) linear scan.
    /// </summary>
    [Params(true, false)]
    public bool UseCache { get; set; }

    [IterationSetup]
    public void IterationSetup()
    {
        CellUtil.UseStyleCache = UseCache;

        // ── HSSF workbook ──────────────────────────────────────────────────
        _hssfWorkbook = new HSSFWorkbook();
        ISheet hssfSheet = _hssfWorkbook.CreateSheet("Sheet1");
        _hssfCells = new ICell[CellCount];
        for (int i = 0; i < CellCount; i++)
        {
            _hssfCells[i] = hssfSheet.CreateRow(i).CreateCell(0);
        }

        // ── XSSF workbook ──────────────────────────────────────────────────
        _xssfWorkbook = new XSSFWorkbook();
        ISheet xssfSheet = _xssfWorkbook.CreateSheet("Sheet1");
        _xssfCells = new ICell[CellCount];
        for (int i = 0; i < CellCount; i++)
        {
            _xssfCells[i] = xssfSheet.CreateRow(i).CreateCell(0);
        }

        // ── Style property maps ────────────────────────────────────────────

        // Single style combo: all cells share the same border style.
        // This is the hot path where caching provides maximum benefit.
        _singleStyleProps = new Dictionary<string, object>
        {
            { CellUtil.BORDER_BOTTOM, BorderStyle.Thin },
            { CellUtil.BORDER_TOP,    BorderStyle.Thin },
            { CellUtil.BORDER_LEFT,   BorderStyle.Thin },
            { CellUtil.BORDER_RIGHT,  BorderStyle.Thin },
        };

        // Multiple distinct style combos (one per cell) to measure cache-miss overhead.
        _multiStyleProps = new Dictionary<string, object>[CellCount];
        for (int i = 0; i < CellCount; i++)
        {
            _multiStyleProps[i] = new Dictionary<string, object>
            {
                { CellUtil.BORDER_BOTTOM, (i % 2 == 0) ? BorderStyle.Thin : BorderStyle.Medium },
                { CellUtil.BORDER_TOP,    (i % 3 == 0) ? BorderStyle.Thin : BorderStyle.Dashed },
                { CellUtil.FILL_FOREGROUND_COLOR, (short)(i % 8) },
            };
        }
    }

    [IterationCleanup]
    public void IterationCleanup()
    {
        CellUtil.UseStyleCache = true; // restore default
        _hssfWorkbook?.Dispose();
        _xssfWorkbook?.Dispose();
    }

    // ── Benchmarks ─────────────────────────────────────────────────────────

    /// <summary>
    /// Applies the same border style to every cell (HSSF).
    /// Best case for cache: only one unique style combo, so after the first
    /// call every subsequent call is a cache hit.
    /// </summary>
    [Benchmark]
    public void HSSF_SingleStyle()
    {
        foreach (ICell cell in _hssfCells)
        {
            CellUtil.SetCellStyleProperties(cell, _singleStyleProps);
        }
    }

    /// <summary>
    /// Applies a different style to each cell (HSSF).
    /// Exercises the cache-miss path; measures whether caching adds overhead
    /// when every style combination is unique.
    /// </summary>
    [Benchmark]
    public void HSSF_MultipleStyles()
    {
        for (int i = 0; i < _hssfCells.Length; i++)
        {
            CellUtil.SetCellStyleProperties(_hssfCells[i], _multiStyleProps[i]);
        }
    }

    /// <summary>
    /// Applies the same border style to every cell (XSSF).
    /// </summary>
    [Benchmark]
    public void XSSF_SingleStyle()
    {
        foreach (ICell cell in _xssfCells)
        {
            CellUtil.SetCellStyleProperties(cell, _singleStyleProps);
        }
    }

    /// <summary>
    /// Applies a different style to each cell (XSSF).
    /// </summary>
    [Benchmark]
    public void XSSF_MultipleStyles()
    {
        for (int i = 0; i < _xssfCells.Length; i++)
        {
            CellUtil.SetCellStyleProperties(_xssfCells[i], _multiStyleProps[i]);
        }
    }
}
