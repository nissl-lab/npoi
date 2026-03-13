using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOI.Benchmarks;

/// <summary>
/// Generates a deterministic .xlsx file with a large <c>xl/styles.xml</c> by
/// creating many unique cell styles and applying them to blank cells.
/// </summary>
internal static class GenerateHugeStylesWorkbook
{
    private const long TargetUncompressedBytes = 30L * 1024 * 1024; // 30 MB
    private const int InitialStyleCount = 10_000;
    private const int StepSize = 5_000;
    private const int MaxStyleCount = 150_000;

    /// <summary>
    /// Creates <paramref name="path"/> if it is missing or its uncompressed
    /// <c>xl/styles.xml</c> is smaller than 30 MB.  Uses step-wise generation:
    /// starts at 10,000 unique styles and increments by 5,000 until the target
    /// is reached or the maximum of 150,000 styles is hit.
    /// </summary>
    public static void EnsureExists(string path)
    {
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        if (File.Exists(path) &&
            ZipUtils.GetStylesXmlUncompressedSize(path) >= TargetUncompressedBytes)
        {
            return;
        }

        int styleCount = InitialStyleCount;
        while (true)
        {
            Generate(path, styleCount);

            if (ZipUtils.GetStylesXmlUncompressedSize(path) >= TargetUncompressedBytes ||
                styleCount >= MaxStyleCount)
            {
                break;
            }

            styleCount = Math.Min(styleCount + StepSize, MaxStyleCount);
        }
    }

    private static void Generate(string path, int styleCount)
    {
        using var wb = new XSSFWorkbook();
        var sheet = wb.CreateSheet("styles");
        var fmt = wb.CreateDataFormat();

        var borderStyles = new[]
        {
            BorderStyle.Thin,
            BorderStyle.Medium,
            BorderStyle.Dashed,
            BorderStyle.Dotted,
        };

        for (int i = 0; i < styleCount; i++)
        {
            var row = sheet.CreateRow(i);
            var cell = row.CreateCell(0);

            var style = wb.CreateCellStyle();

            // Unique number format per cell
            style.DataFormat = fmt.GetFormat($"0.{new string('0', (i % 6) + 1)}E+{i:D4}");

            // Unique font variation
            var font = wb.CreateFont();
            font.FontName = (i % 2 == 0) ? "Calibri" : "Arial";
            font.FontHeightInPoints = (short)(8 + (i % 8));
            font.IsBold = (i % 3 == 0);
            font.IsItalic = (i % 5 == 0);
            font.Color = (short)(8 + (i % 56));
            style.SetFont(font);

            // Unique fill
            style.FillPattern = FillPattern.SolidForeground;
            style.FillForegroundColor = (short)(8 + ((i * 7) % 56));

            // Unique border
            var bs = borderStyles[i % borderStyles.Length];
            style.BorderBottom = bs;
            style.BorderTop = bs;
            style.BorderLeft = bs;
            style.BorderRight = bs;

            cell.CellStyle = style;
        }

        using var fs = File.Create(path);
        wb.Write(fs);
    }
}
