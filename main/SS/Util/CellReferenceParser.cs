using System;

namespace NPOI.SS.Util;

internal static class CellReferenceParser
{
    // (\$?[A-Za-z]+)?(\$?[0-9]+)?
    public static bool TryParseCellReference(ReadOnlySpan<char> input, out char columnPrefix, out ReadOnlySpan<char> column, out char rowPrefix, out ReadOnlySpan<char> row)
    {
        return TryParse(input, out columnPrefix, out column, out rowPrefix, out row)
               && columnPrefix is '$' or char.MinValue && rowPrefix is '$' or char.MinValue;
    }

    // Matches references only where row and column are included.
    // Matches a run of one or more letters followed by a run of one or more digits.
    // If a reference does not match this pattern, it might match COLUMN_REF_PATTERN or ROW_REF_PATTERN
    // References may optionally include a single '$' before each group, but these are excluded from the Matcher.group(int).
    // ^\$?([A-Z]+)\$?([0-9]+)$
    public static bool TryParseStrictCellReference(ReadOnlySpan<char> input, out ReadOnlySpan<char> column, out ReadOnlySpan<char> row)
    {
        return TryParse(input, out var columnPrefix, out column, out var rowPrefix, out row)
               && columnPrefix is '$' or char.MinValue
               && column.Length > 0
               && rowPrefix is '$' or char.MinValue
               && row.Length > 0;
    }


    // Matches a run of one or more letters.  The run of letters is group 1.
    // References may optionally include a single '$' before the group, but these are excluded from the Matcher.group(int).
    // ^\$?([A-Za-z]+)$
    public static bool TryParseColumnReference(ReadOnlySpan<char> input, out ReadOnlySpan<char> column)
    {
        return TryParse(input, out var columnPrefix, out column, out var rowPrefix, out var row)
               && columnPrefix is '$' or char.MinValue
               && column.Length > 0
               && rowPrefix is char.MinValue
               && row.Length == 0;
    }

    // Matches a run of one or more letters.  The run of numbers is group 1.
    // References may optionally include a single '$' before the group, but these are excluded from the Matcher.group(int).
    // ^\$?([0-9]+)$
    public static bool TryParseRowReference(ReadOnlySpan<char> input, out ReadOnlySpan<char> row)
    {
        return TryParse(input, out var columnPrefix, out var cell, out var rowPrefix, out row)
               && columnPrefix is '$' or char.MinValue
               && cell.Length == 0
               && rowPrefix is '$' or char.MinValue
               && row.Length > 0;
    }

    /// <summary>
    /// Generic parsing logic that extracts reference information.
    /// </summary>
    /// <param name="input">Input to parse.</param>
    /// <param name="columnPrefix">Possible column prefix like '$', <see cref="char.MinValue" /> if none.</param>
    /// <param name="column">Column name string, empty if none.</param>
    /// <param name="rowPrefix">Possible row prefix like '$', <see cref="char.MinValue" /> if none.</param>
    /// <param name="row">Row data, empty if none</param>
    /// <returns></returns>
    private static bool TryParse(
        ReadOnlySpan<char> input,
        out char columnPrefix,
        out ReadOnlySpan<char> column,
        out char rowPrefix,
        out ReadOnlySpan<char> row)
    {
        column = default;
        columnPrefix = char.MinValue;
        row = default;
        rowPrefix = char.MinValue;

        if (input.Length == 0)
        {
            return false;
        }

        // quick check for common case, alphabet + numbers, like A11
        var firstChar = input[0];
        if (input.Length > 1
            && char.IsLetter(firstChar)
            && char.IsDigit(input[1])
            && TryParsePositiveInt32Fast(input.Slice(1), out _))
        {
            column = input.Slice(0, 1);
            row = input.Slice(1);
            return true;
        }

        int cellStartIndex = 0;
        int cellEndIndex = input.Length - 1;
        int rowStartIndex = input.Length;

        if (char.IsDigit(firstChar))
        {
            // no cell
            cellStartIndex = int.MaxValue;
            rowStartIndex = 0;
        }
        else if (!char.IsLetter(firstChar))
        {
            if (input.Length > 1 && char.IsDigit(input[1]))
            {
                // actually row starts now
                rowStartIndex = 0;
                cellStartIndex = input.Length;
            }
            else
            {
                columnPrefix = firstChar;
                cellStartIndex = 1;
            }
        }

        for (int i = cellStartIndex; i < input.Length; ++i)
        {
            var c = input[i];
            cellEndIndex = i + 1;
            if (!char.IsLetter(c))
            {
                // end of cell information
                rowStartIndex = i;
                cellEndIndex--;
                break;
            }
        }

        for (int i = rowStartIndex; i < input.Length; ++i)
        {
            var c = input[i];

            if (!char.IsNumber(c) && i == rowStartIndex)
            {
                // first is allowed to be a prefix
                rowPrefix = c;
                rowStartIndex++;
                continue;
            }

            if (!char.IsDigit(input[i]))
            {
                return false;
            }
        }

        // seems ok
        var cellStringLength = cellEndIndex - cellStartIndex;
        if (cellStringLength > 0)
        {
            column = input.Slice(cellStartIndex, cellStringLength);
        }

        row = input.Slice(rowStartIndex);
        return true;
    }

    public static bool TryParsePositiveInt32Fast(ReadOnlySpan<char> s, out int result)
    {
        int value = 0;
        foreach (var c in s)
        {
            if (!char.IsDigit(c))
            {
                result = -1;
                return false;
            }

            value = 10 * value + (c - 48);
        }

        result = value;
        return true;
    }
}