using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOI.Benchmarks;

/// <summary>
/// Generates a deterministic .xlsx file with a large <c>xl/styles.xml</c> by
/// creating many unique cell styles and applying them to blank cells.
///
/// <para>
/// Each style carries a unique font (mixed-radix height/colour/bold/italic/name
/// combination), a varied alignment sub-element, and extra font XML attributes
/// (underline, strikeout, charset, family, scheme, shadow, condense, extend)
/// that increase the bytes-per-entry in the serialised styles.xml to ~470 bytes
/// per style.  With NPOI's 64 000-style xlsx ceiling the practical maximum
/// uncompressed styles.xml size is ~28 MB; a 20 MB target therefore lands
/// comfortably in the 20–28 MB range.
/// </para>
/// </summary>
internal static class GenerateHugeStylesWorkbook
{
    // Target ~20 MB uncompressed styles.xml (achievable within NPOI's 64k-style
    // xlsx cap; the loop stops around 45k styles producing ~20 MB).
    private const long TargetUncompressedBytes = 20L * 1024 * 1024;
    private const int InitialStyleCount = 10_000;
    private const int StepSize = 5_000;
    // NPOI enforces a 64,000-style limit (xlsx spec theoretical max: 65,530).
    private const int MaxStyleCount = 63_500;

    /// <summary>
    /// Creates <paramref name="path"/> if it is missing or its uncompressed
    /// <c>xl/styles.xml</c> is smaller than 20 MB.  Uses step-wise generation:
    /// starts at 10,000 unique styles and increments by 5,000 until the target
    /// is reached or the xlsx-spec maximum of 63,500 styles is hit.
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

        // Pre-build a fixed set of ≤20 number formats so we stay well under the
        // 250-custom-format limit while still contributing varied numFmtId values.
        const int NumFmtCount = 20;
        var numFmtIds = new short[NumFmtCount];
        for (int f = 0; f < NumFmtCount; f++)
            numFmtIds[f] = fmt.GetFormat($"0.{new string('0', f + 1)}");

        // Four border styles (small fixed set — keeps border dedup O(constant)).
        var borderStyles = new[]
        {
            BorderStyle.Thin,
            BorderStyle.Medium,
            BorderStyle.Dashed,
            BorderStyle.Dotted,
        };

        // Mixed-radix dimensions that give 200 × 56 × 2 × 2 × 2 = 179,200 unique
        // font combinations — more than enough for the 63,500 style ceiling.
        // XSSFWorkbook.CreateFont() uses forceRegistration=true so font creation
        // is O(1) with no dedup scan, keeping generation time linear in styleCount.
        const int HeightRange = 200; // point sizes 8–207
        const int ColorRange  = 56;  // indexed colours 8–63

        for (int i = 0; i < styleCount; i++)
        {
            var row = sheet.CreateRow(i);
            var cell = row.CreateCell(0);

            var style = (XSSFCellStyle)wb.CreateCellStyle();
            style.DataFormat = numFmtIds[i % NumFmtCount];

            // Decompose i into independent font dimensions (mixed-radix).
            // Every (height, colour, bold, italic, name) tuple is unique for
            // i < 179,200, ensuring a unique font entry per style.
            int fi = i;
            var font = (XSSFFont)wb.CreateFont();
            font.FontHeightInPoints = (short)(8 + (fi % HeightRange));
            fi /= HeightRange;
            font.Color    = (short)(8 + (fi % ColorRange));
            fi /= ColorRange;
            font.IsBold   = (fi % 2 == 0);
            fi /= 2;
            font.IsItalic = (fi % 2 == 0);
            fi /= 2;
            font.FontName = (fi % 2 == 0) ? "Calibri" : "Arial";

            // Verbose font attributes: each non-default value is serialised to XML,
            // increasing each <font> entry's byte count.
            font.Underline   = (fi % 2 == 0) ? FontUnderlineType.Single : FontUnderlineType.Double;
            font.IsStrikeout = (i % 3 == 0);
            font.Charset     = (short)(i % 3 == 0 ? 161 : 0);
            font.TypeOffset  = (FontSuperScript)(i % 3);    // adds <vertAlign>
            font.Family      = 2;                            // adds <family val="2"/>
            font.SetScheme(FontScheme.MINOR);                // adds <scheme val="minor"/>

            // shadow/condense/extend are not exposed on IFont; access via CT_Font.
            var ctFont = font.GetCTFont();
            ctFont.AddNewShadow().val   = (i % 2 == 0);     // adds <shadow val="..."/>
            ctFont.AddNewCondense().val = (i % 3 == 0);     // adds <condense val="..."/>
            ctFont.AddNewExtend().val   = (i % 5 == 0);     // adds <extend val="..."/>
            style.SetFont(font);

            // Fill: cycle through 56 distinct indexed foreground colours.
            style.FillPattern = FillPattern.SolidForeground;
            style.FillForegroundColor = (short)(8 + (i % ColorRange));

            // Border: cycle through the 4 border styles (tiny fixed set = fast dedup).
            var bs = borderStyles[i % borderStyles.Length];
            style.BorderBottom = bs;
            style.BorderTop    = bs;
            style.BorderLeft   = bs;
            style.BorderRight  = bs;

            // Verbose alignment sub-element: adds ~130 bytes to each <xf> entry.
            style.Alignment         = (i % 2 == 0) ? HorizontalAlignment.Left : HorizontalAlignment.Right;
            style.VerticalAlignment = (i % 3 == 0) ? VerticalAlignment.Top    : VerticalAlignment.Center;
            style.WrapText          = (i % 2 == 0);
            style.Indention         = (short)(i % 8);
            style.Rotation          = (short)(i % 90);
            style.ShrinkToFit       = (i % 5 == 0);
            style.ReadingOrder      = (i % 2 == 0) ? ReadingOrder.LEFT_TO_RIGHT : ReadingOrder.RIGHT_TO_LEFT;
            style.IsQuotePrefixed   = (i % 3 == 0);
            style.IsLocked          = true;
            style.IsHidden          = (i % 7 == 0);

            cell.CellStyle = style;
        }

        using var fs = File.Create(path);
        wb.Write(fs);
    }
}


