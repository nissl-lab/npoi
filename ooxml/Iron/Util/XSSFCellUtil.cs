using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NPOI.XSSF.Util
{
    public class XSSFCellUtil
    {
        public const string ALIGNMENT = "alignment";
        public const string BORDER_BOTTOM = "borderBottom";
        public const string BORDER_DIAGONAL = "borderDiagonal";
        public const string BORDER_LEFT = "borderLeft";
        public const string BORDER_RIGHT = "borderRight";
        public const string BORDER_TOP = "borderTop";
        public const string BOTTOM_BORDER_COLOR = "bottomBorderColor";
        public const string DATA_FORMAT = "dataFormat";
        public const string FILL_BACKGROUND_COLOR = "fillBackgroundColor";
        public const string FILL_FOREGROUND_COLOR = "fillForegroundColor";
        public const string FILL_PATTERN = "fillPattern";
        public const string FONT = "font";
        public const string HIDDEN = "hidden";
        public const string INDENTION = "indention";
        public const string LEFT_BORDER_COLOR = "leftBorderColor";
        public const string LOCKED = "locked";
        public const string RIGHT_BORDER_COLOR = "rightBorderColor";
        public const string ROTATION = "rotation";
        public const string SHRINK_TO_FIT = "shrinkToFit";
        public const string TOP_BORDER_COLOR = "topBorderColor";
        public const string VERTICAL_ALIGNMENT = "verticalAlignment";
        public const string WRAP_TEXT = "wrapText";

        private static readonly ISet<string> shortValues = new HashSet<string>(new string[]{
            INDENTION,
            DATA_FORMAT,
            ROTATION,
        });
        private static readonly ISet<string> intValues = new HashSet<string>(new string[]{
            FONT,
        });
        private static readonly ISet<string> booleanValues = new HashSet<string>(new string[]{
            LOCKED,
            HIDDEN,
            WRAP_TEXT,
            SHRINK_TO_FIT,
        });
        private static readonly ISet<string> borderTypeValues = new HashSet<string>(new string[]{
            BORDER_BOTTOM,
            BORDER_LEFT,
            BORDER_RIGHT,
            BORDER_TOP,
        });
        private static readonly ISet<string> stringValues = new HashSet<string>(new string[]
        {
            BOTTOM_BORDER_COLOR,
            LEFT_BORDER_COLOR,
            RIGHT_BORDER_COLOR,
            TOP_BORDER_COLOR,
            FILL_FOREGROUND_COLOR,
            FILL_BACKGROUND_COLOR,
        });

        private XSSFCellUtil()
        {
        }

        public static void SetFont(ICell cell, XSSFWorkbook workbook, IFont font)
        {
            SetCellStyleProperty(cell, workbook, FONT, (int)font.Index);
        }

        public static void SetCellStyleProperty(ICell cell, XSSFWorkbook workbook, string propertyName, object propertyValue)
        {
            SetCellStyleProperties(cell, workbook, new Dictionary<string, object>() { { propertyName, propertyValue } });
        }

        public static void SetCellStyleProperties(ICell cell, XSSFWorkbook workbook, Dictionary<string, object> propertiesToSet)
        {
            var styleMap = GetFormatProperties(cell.CellStyle as XSSFCellStyle);
            PutAll(propertiesToSet, styleMap);

            var style = LookUpOrCreateCellStyleInWorkbook(styleMap, workbook);

            cell.CellStyle = style;
        }

        public static ICellStyle LookUpOrCreateCellStyleInWorkbook(XSSFCellStyle lookUpStyle, XSSFWorkbook workbook)
        {
            var styleMap = GetFormatProperties(lookUpStyle);

            return LookUpOrCreateCellStyleInWorkbook(styleMap, workbook);
        }

        public static ICellStyle LookUpOrCreateCellStyleInWorkbook(Dictionary<string, object> lookUpStyleMap, XSSFWorkbook workbook)
        {
            for (var i = 0; i < workbook.NumCellStyles; i++)
            {
                var workbookStyle = workbook.GetCellStyleAt(i);
                var workbookStyleMap = GetFormatProperties(workbookStyle as XSSFCellStyle);

                var stylesAreEqual = CompareStyleMaps(lookUpStyleMap, workbookStyleMap);

                if (stylesAreEqual)
                {
                    return workbookStyle;
                }
            }

            var newStyle = workbook.CreateCellStyle();
            SetFormatProperties(newStyle as XSSFCellStyle, workbook, lookUpStyleMap);

            return newStyle;
        }

        public static bool CompareStyles(XSSFCellStyle styleA, XSSFCellStyle styleB)
        {
            var styleMapA = GetFormatProperties(styleA);
            var styleMapB = GetFormatProperties(styleB);

            return CompareStyleMaps(styleMapA, styleMapB);
        }

        private static bool CompareStyleMaps(Dictionary<string, object> x, Dictionary<string, object> y)
        {
            if (x.Count != y.Count)
            {
                return false;
            }

            if (x.Keys.Except(y.Keys).Any())
            {
                return false;
            }

            if (y.Keys.Except(x.Keys).Any())
            {
                return false;
            }

            foreach (var key in x.Keys)
            {
                if (ValuesAreNotEqual(x[key], y[key]))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool ValuesAreNotEqual(object xValue, object yValue)
        {
            return (xValue != null && yValue == null) ||
                (xValue == null && yValue != null) ||
                (xValue != null && yValue != null && !xValue.Equals(yValue));
        }

        private static void PutAll(Dictionary<string, object> src, Dictionary<string, object> dest)
        {
            foreach (var key in src.Keys)
            {
                if (shortValues.Contains(key))
                {
                    dest[key] = GetShort(src, key);
                }
                else if (intValues.Contains(key))
                {
                    dest[key] = GetInt(src, key);
                }
                else if (stringValues.Contains(key))
                {
                    dest[key] = GetString(src, key);
                }
                else if (booleanValues.Contains(key))
                {
                    dest[key] = GetBoolean(src, key);
                }
                else if (borderTypeValues.Contains(key))
                {
                    dest[key] = GetBorderStyle(src, key);
                }
                else if (ALIGNMENT.Equals(key))
                {
                    dest[key] = GetHorizontalAlignment(src, key);
                }
                else if (VERTICAL_ALIGNMENT.Equals(key))
                {
                    dest[key] = GetVerticalAlignment(src, key);
                }
                else if (FILL_PATTERN.Equals(key))
                {
                    dest[key] = GetFillPattern(src, key);
                }
            }
        }

        public static Dictionary<string, object> GetFormatProperties(XSSFCellStyle style)
        {
            var properties = new Dictionary<string, object>();
            Put(properties, ALIGNMENT, style.Alignment);
            Put(properties, BORDER_BOTTOM, style.BorderBottom);
            PutShort(properties, BORDER_DIAGONAL, (short)style.BorderDiagonal);
            Put(properties, BORDER_LEFT, style.BorderLeft);
            Put(properties, BORDER_RIGHT, style.BorderRight);
            Put(properties, BORDER_TOP, style.BorderTop);
            PutString(properties, BOTTOM_BORDER_COLOR, RgbByteArrayToHexstring(style.BottomBorderXSSFColor?.GetRgbWithTint() ?? style.BottomBorderXSSFColor?.RGB));
            PutShort(properties, DATA_FORMAT, style.DataFormat);
            PutString(properties, FILL_BACKGROUND_COLOR, RgbByteArrayToHexstring(style.FillBackgroundXSSFColor?.GetRgbWithTint() ?? style.FillBackgroundXSSFColor?.RGB));
            PutString(properties, FILL_FOREGROUND_COLOR, RgbByteArrayToHexstring(style.FillForegroundXSSFColor?.GetRgbWithTint() ?? style.FillForegroundXSSFColor?.RGB));
            Put(properties, FILL_PATTERN, style.FillPattern);
            PutInt(properties, FONT, style.FontIndex);
            PutBoolean(properties, HIDDEN, style.IsHidden);
            PutShort(properties, INDENTION, style.Indention);
            PutString(properties, LEFT_BORDER_COLOR, RgbByteArrayToHexstring(style.LeftBorderXSSFColor?.GetRgbWithTint() ?? style.LeftBorderXSSFColor?.RGB));
            PutBoolean(properties, LOCKED, style.IsLocked);
            PutString(properties, RIGHT_BORDER_COLOR, RgbByteArrayToHexstring(style.RightBorderXSSFColor?.GetRgbWithTint() ?? style.RightBorderXSSFColor?.RGB));
            PutShort(properties, ROTATION, style.Rotation);
            PutBoolean(properties, SHRINK_TO_FIT, style.ShrinkToFit);
            PutString(properties, TOP_BORDER_COLOR, RgbByteArrayToHexstring(style.TopBorderXSSFColor?.GetRgbWithTint() ?? style.TopBorderXSSFColor?.RGB));
            Put(properties, VERTICAL_ALIGNMENT, style.VerticalAlignment);
            PutBoolean(properties, WRAP_TEXT, style.WrapText);
            return properties;
        }
        
        private static void Put(Dictionary<string, object> properties, string name, object value)
        {
            properties[name] = value;
        }

        private static void SetFormatProperties(XSSFCellStyle style, XSSFWorkbook workbook, Dictionary<string, object> properties)
        {
            style.Alignment = GetHorizontalAlignment(properties, ALIGNMENT);
            style.BorderBottom = GetBorderStyle(properties, BORDER_BOTTOM);

            style.BorderDiagonal = (BorderDiagonal)GetShort(properties, BORDER_DIAGONAL);

            style.BorderLeft = GetBorderStyle(properties, BORDER_LEFT);
            style.BorderRight = GetBorderStyle(properties, BORDER_RIGHT);
            style.BorderTop = GetBorderStyle(properties, BORDER_TOP);

            var bottomBorderColor = StringToByteArray(GetString(properties, BOTTOM_BORDER_COLOR));
            if (bottomBorderColor != null)
            {
                style.SetBottomBorderColor(new XSSFColor(bottomBorderColor));
            }

            style.DataFormat = GetShort(properties, DATA_FORMAT);

            var backgroundColor = StringToByteArray(GetString(properties, FILL_BACKGROUND_COLOR));
            if (backgroundColor != null)
            {
                style.SetFillBackgroundColor(new XSSFColor(backgroundColor));
            }

            var foregroundColor = StringToByteArray(GetString(properties, FILL_FOREGROUND_COLOR));
            if (foregroundColor != null)
            {
                style.SetFillForegroundColor(new XSSFColor(foregroundColor));
            }

            style.FillPattern = GetFillPattern(properties, FILL_PATTERN);
            style.SetFont(workbook.GetFontAt(GetShort(properties, FONT)));
            style.IsHidden = GetBoolean(properties, HIDDEN);
            style.Indention = GetShort(properties, INDENTION);

            var leftBorderColor = StringToByteArray(GetString(properties, LEFT_BORDER_COLOR));
            if (leftBorderColor != null)
            {
                style.SetLeftBorderColor(new XSSFColor(leftBorderColor));
            }

            style.IsLocked = GetBoolean(properties, LOCKED);

            var rightBorderColor = StringToByteArray(GetString(properties, RIGHT_BORDER_COLOR));
            if (rightBorderColor != null)
            {
                style.SetRightBorderColor(new XSSFColor(rightBorderColor));
            }

            style.Rotation = GetShort(properties, ROTATION);
            style.ShrinkToFit = GetBoolean(properties, SHRINK_TO_FIT);

            var topBorderColor = StringToByteArray(GetString(properties, TOP_BORDER_COLOR));
            if (topBorderColor != null)
            {
                style.SetTopBorderColor(new XSSFColor(topBorderColor));
            }

            style.VerticalAlignment = GetVerticalAlignment(properties, VERTICAL_ALIGNMENT);
            style.WrapText = GetBoolean(properties, WRAP_TEXT);
        }

        private static short GetShort(Dictionary<string, object> properties, string name)
        {
            var value = properties[name];
            return short.TryParse(value.ToString(), out var result) ? result : (short)0;
        }

        private static int GetInt(Dictionary<string, object> properties, string name)
        {
            var value = properties[name];
            if (Number.IsNumber(value))
            {
                return int.Parse(value.ToString());
            }

            return 0;
        }

        private static BorderStyle GetBorderStyle(Dictionary<string, object> properties, string name)
        {
            var value = properties[name];
            BorderStyle border;
            if (value is BorderStyle borderStyle)
            {
                border = borderStyle;
            }
            else if (value is short || value is int)
            {
                var code = short.Parse(value.ToString());
                border = (BorderStyle)code;
            }
            else if (value == null)
            {
                border = BorderStyle.None;
            }
            else
            {
                throw new RuntimeException("Unexpected border style class. Must be BorderStyle or Short (deprecated).");
            }

            return border;
        }

        private static FillPattern GetFillPattern(Dictionary<string, object> properties, string name)
        {
            var value = properties[name];
            FillPattern pattern;
            if (value is FillPattern ptrn)
            {
                pattern = ptrn;
            }
            else if (value is short code)
            {
                pattern = (FillPattern)code;
            }
            else if (value == null)
            {
                pattern = FillPattern.NoFill;
            }
            else
            {
                throw new RuntimeException("Unexpected fill pattern style class. Must be FillPattern or Short (deprecated).");
            }

            return pattern;
        }

        private static HorizontalAlignment GetHorizontalAlignment(Dictionary<string, object> properties, string name)
        {
            var value = properties[name];
            HorizontalAlignment align;
            if (value is HorizontalAlignment alignment)
            {
                align = alignment;
            }
            else if (value is short code)
            {
                align = (HorizontalAlignment)code;
            }
            else if (value == null)
            {
                align = HorizontalAlignment.General;
            }
            else
            {
                throw new RuntimeException("Unexpected horizontal alignment style class. Must be HorizontalAlignment or Short (deprecated).");
            }

            return align;
        }

        private static VerticalAlignment GetVerticalAlignment(Dictionary<string, object> properties, string name)
        {
            var value = properties[name];
            VerticalAlignment align;
            if (value is VerticalAlignment alignment)
            {
                align = alignment;
            }
            else if (value is short code)
            {
                align = (VerticalAlignment)code;
            }
            else if (value == null)
            {
                align = VerticalAlignment.Bottom;
            }
            else
            {
                throw new RuntimeException("Unexpected vertical alignment style class. Must be VerticalAlignment or Short (deprecated).");
            }

            return align;
        }

        private static bool GetBoolean(Dictionary<string, object> properties, string name)
        {
            var value = properties[name];

            if (bool.TryParse(value.ToString(), out var result))
            {
                return result;
            }

            return false;
        }
        private static string GetString(Dictionary<string, object> properties, string name)
        {
            var value = properties[name];
            return value?.ToString();
        }

        private static void PutShort(Dictionary<string, object> properties, string name, short value)
        {
            if (properties.ContainsKey(name))
            {
                properties[name] = value;
            }
            else
            {
                properties.Add(name, value);
            }
        }

        private static void PutInt(Dictionary<string, object> properties, string name, int value)
        {
            if (properties.ContainsKey(name))
            {
                properties[name] = value;
            }
            else
            {
                properties.Add(name, value);
            }
        }

        private static void PutString(Dictionary<string, object> properties, string name, string value)
        {
            if (properties.ContainsKey(name))
            {
                properties[name] = value;
            }
            else
            {
                properties.Add(name, value);
            }
        }

        private static void PutBoolean(Dictionary<string, object> properties, string name, bool value)
        {
            if (properties.ContainsKey(name))
            {
                properties[name] = value;
            }
            else
            {
                properties.Add(name, value);
            }
        }

        private static string RgbByteArrayToHexstring(byte[] rgb)
        {
            if (rgb == null)
            {
                return null;
            }

            return "#" + string.Format("{0:x2}{1:x2}{2:x2}", rgb[0], rgb[1], rgb[2]);
        }

        private static byte[] StringToByteArray(string hexString)
        {
            if (hexString == null)
            {
                return null;
            }

            if (hexString.StartsWith("#"))
            {
                hexString = hexString.Substring(1);
            }

            return Enumerable.Range(0, hexString.Length)
                .Where(x => x % 2 == 0)
                .Select(x => Convert.ToByte(hexString.Substring(x, 2), 16))
                .ToArray();
        }
    }
}
