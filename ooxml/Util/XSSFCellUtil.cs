using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NPOI.XSSF.Util
{
    //For XLSX files we need to copy and find colors by rgb string, not by color index.
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
        public const string DIAGONAL_BORDER_COLOR = "diagonalBorderColor";
        public const string DIAGONAL_BORDER_LINE_STYLE = "diagonalBorderLineStyle";
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

        private XSSFCellUtil()
        {
            // no instances of this class
        }

        public static void SetFont(ICell cell, XSSFWorkbook workbook, IFont font)
        {
            SetCellStyleProperty(cell, workbook, FONT, font.Index);
        }

        public static void SetCellStyleProperty(ICell cell, XSSFWorkbook workbook, string propertyName, object propertyValue)
        {
            SetCellStyleProperties(cell, workbook, new List<string> { propertyName }, new List<object> { propertyValue });
        }

        public static void SetCellStyleProperties(ICell cell, XSSFWorkbook workbook, List<string> propertyNames, List<object> propertyValues)
        {
            ICellStyle originalStyle = cell.CellStyle;
            ICellStyle newStyle = null;
            Dictionary<string, object> values = GetFormatProperties(originalStyle as XSSFCellStyle);
            if (propertyNames.Count != propertyValues.Count)
            {
                throw new ArgumentException("Amount of names and properties should be equal");
            }

            for (var i = 0; i < propertyNames.Count; i++)
            {
                var propertyName = propertyNames[i];
                var propertyValue = propertyValues[i];
                values[propertyName] = (propertyValue is Enum) ? (short)propertyValue : propertyValue;
            }

            int numberCellStyles = workbook.NumCellStyles;

            for (short i = 0; i < numberCellStyles; i++)
            {
                ICellStyle wbStyle = workbook.GetCellStyleAt(i);

                Dictionary<string, object> wbStyleMap = GetFormatProperties(wbStyle as XSSFCellStyle);

                if (values.Keys.Count != wbStyleMap.Keys.Count)
                {
                    continue;
                }

                bool found = true;

                foreach (string key in values.Keys)
                {
                    if (!wbStyleMap.ContainsKey(key))
                    {
                        found = false;
                        break;
                    }

                    if (values[key] == null && wbStyleMap[key] == null || values[key] != null && values[key].Equals(wbStyleMap[key]))
                    {
                        continue;
                    }

                    found = false;
                    break;
                }

                if (found)
                {
                    newStyle = wbStyle;
                    break;
                }
            }

            if (newStyle == null)
            {
                newStyle = workbook.CreateCellStyle();
                SetFormatProperties(newStyle as XSSFCellStyle, workbook, values);
            }

            cell.CellStyle = newStyle;
        }

        private static Dictionary<string, object> GetFormatProperties(XSSFCellStyle style)
        {
            Dictionary<string, object> properties = new();
            PutShort(properties, ALIGNMENT, (short)style.Alignment);
            PutShort(properties, BORDER_BOTTOM, (short)style.BorderBottom);
            PutShort(properties, BORDER_DIAGONAL, (short)style.BorderDiagonal);
            PutShort(properties, BORDER_LEFT, (short)style.BorderLeft);
            PutShort(properties, BORDER_RIGHT, (short)style.BorderRight);
            PutShort(properties, BORDER_TOP, (short)style.BorderTop);
            PutString(properties, BOTTOM_BORDER_COLOR, RgbByteArrayToHexstring(style.BottomBorderXSSFColor?.GetRgbWithTint() ?? style.BottomBorderXSSFColor?.RGB));
            PutShort(properties, DATA_FORMAT, style.DataFormat);
            PutString(properties, DIAGONAL_BORDER_COLOR, RgbByteArrayToHexstring(style.DiagonalBorderXSSFColor?.GetRgbWithTint() ?? style.DiagonalBorderXSSFColor?.RGB));
            PutShort(properties, DIAGONAL_BORDER_LINE_STYLE, (short)style.BorderDiagonalLineStyle);
            PutString(properties, FILL_BACKGROUND_COLOR, RgbByteArrayToHexstring(style.FillBackgroundXSSFColor?.GetRgbWithTint() ?? style.FillBackgroundXSSFColor?.RGB));
            PutString(properties, FILL_FOREGROUND_COLOR, RgbByteArrayToHexstring(style.FillForegroundXSSFColor?.GetRgbWithTint() ?? style.FillForegroundXSSFColor?.RGB));
            PutShort(properties, FILL_PATTERN, (short)style.FillPattern);
            PutShort(properties, FONT, style.FontIndex);
            PutBoolean(properties, HIDDEN, style.IsHidden);
            PutShort(properties, INDENTION, style.Indention);
            PutString(properties, LEFT_BORDER_COLOR, RgbByteArrayToHexstring(style.LeftBorderXSSFColor?.GetRgbWithTint() ?? style.LeftBorderXSSFColor?.RGB));
            PutBoolean(properties, LOCKED, style.IsLocked);
            PutString(properties, RIGHT_BORDER_COLOR, RgbByteArrayToHexstring(style.RightBorderXSSFColor?.GetRgbWithTint() ?? style.RightBorderXSSFColor?.RGB));
            PutShort(properties, ROTATION, style.Rotation);
            PutBoolean(properties, SHRINK_TO_FIT, style.ShrinkToFit);
            PutString(properties, TOP_BORDER_COLOR, RgbByteArrayToHexstring(style.TopBorderXSSFColor?.GetRgbWithTint() ?? style.TopBorderXSSFColor?.RGB));
            PutShort(properties, VERTICAL_ALIGNMENT, (short)style.VerticalAlignment);
            PutBoolean(properties, WRAP_TEXT, style.WrapText);
            return properties;
        }

        private static void SetFormatProperties(XSSFCellStyle style, XSSFWorkbook workbook, Dictionary<string, object> properties)
        {
            style.Alignment = (HorizontalAlignment)GetShort(properties, ALIGNMENT);
            style.BorderBottom = (BorderStyle)GetShort(properties, BORDER_BOTTOM);

            style.BorderDiagonal = (BorderDiagonal)GetShort(properties, BORDER_DIAGONAL);
            style.BorderDiagonalLineStyle = (BorderStyle)GetShort(properties, DIAGONAL_BORDER_LINE_STYLE);

            var diagonalBorderColor = StringToByteArray(GetString(properties, DIAGONAL_BORDER_COLOR));
            if (diagonalBorderColor != null)
            {
                style.SetDiagonalBorderColor(new XSSFColor(diagonalBorderColor));
            }
            else
            {
                style.BorderDiagonalColor = 0;
            }

            style.BorderLeft = (BorderStyle)GetShort(properties, BORDER_LEFT);
            style.BorderRight = (BorderStyle)GetShort(properties, BORDER_RIGHT);
            style.BorderTop = (BorderStyle)GetShort(properties, BORDER_TOP);

            var bottomBorderColor = StringToByteArray(GetString(properties, BOTTOM_BORDER_COLOR));
            if (bottomBorderColor != null)
            {
                style.SetBottomBorderColor(new XSSFColor(bottomBorderColor));
            }
            else
            {
                style.BottomBorderColor = 0;
            }

            style.DataFormat = GetShort(properties, DATA_FORMAT);

            var backgroundColor = StringToByteArray(GetString(properties, FILL_BACKGROUND_COLOR));
            if (backgroundColor != null)
            {
                style.SetFillBackgroundColor(new XSSFColor(backgroundColor));
            }
            else
            {
                style.FillBackgroundColor = 0;
            }

            var foregroundColor = StringToByteArray(GetString(properties, FILL_FOREGROUND_COLOR));
            if (foregroundColor != null)
            {
                style.SetFillForegroundColor(new XSSFColor(foregroundColor));
            }
            else
            {
                style.FillForegroundColor = 0;
            }

            style.FillPattern = (FillPattern)GetShort(properties, FILL_PATTERN);
            style.SetFont(workbook.GetFontAt(GetShort(properties, FONT)));
            style.IsHidden = GetBoolean(properties, HIDDEN);
            style.Indention = GetShort(properties, INDENTION);

            var leftBorderColor = StringToByteArray(GetString(properties, LEFT_BORDER_COLOR));
            if (leftBorderColor != null)
            {
                style.SetLeftBorderColor(new XSSFColor(leftBorderColor));
            }
            else
            {
                style.LeftBorderColor = 0;
            }

            style.IsLocked = GetBoolean(properties, LOCKED);

            var rightBorderColor = StringToByteArray(GetString(properties, RIGHT_BORDER_COLOR));
            if (rightBorderColor != null)
            {
                style.SetRightBorderColor(new XSSFColor(rightBorderColor));
            }
            else
            {
                style.RightBorderColor = 0;
            }

            style.Rotation = GetShort(properties, ROTATION);
            style.ShrinkToFit = GetBoolean(properties, SHRINK_TO_FIT);

            var topBorderColor = StringToByteArray(GetString(properties, TOP_BORDER_COLOR));
            if (topBorderColor != null)
            {
                style.SetTopBorderColor(new XSSFColor(topBorderColor));
            }
            else
            {
                style.TopBorderColor = 0;
            }

            style.VerticalAlignment = (VerticalAlignment)GetShort(properties, VERTICAL_ALIGNMENT);
            style.WrapText = GetBoolean(properties, WRAP_TEXT);
        }

        /**
         * Utility method that returns the named short value form the given map.
         * @return zero if the property does not exist, or is not a {@link Short}.
         *
         * @param properties map of named properties (String -> Object)
         * @param name property name
         * @return property value, or zero
         */
        private static short GetShort(Dictionary<string, object> properties, string name)
        {
            object value = properties[name];

            if (short.TryParse(value.ToString(), out short result))
            {
                return result;
            }

            return (short)(int)value;
        }

        private static string GetString(Dictionary<string, object> properties, string name)
        {
            object value = properties[name];
            return value?.ToString();
        }

        /**
         * Utility method that returns the named boolean value form the given map.
         * @return false if the property does not exist, or is not a {@link Boolean}.
         *
         * @param properties map of properties (String -> Object)
         * @param name property name
         * @return property value, or false
         */
        private static bool GetBoolean(Dictionary<string, object> properties, string name)
        {
            object value = properties[name];

            if (bool.TryParse(value.ToString(), out bool result))
            {
                return result;
            }

            return false;
        }

        /**
         * Utility method that puts the named short value to the given map.
         *
         * @param properties map of properties (String -> Object)
         * @param name property name
         * @param value property value
         */
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

        /**
         * Utility method that puts the named boolean value to the given map.
         *
         * @param properties map of properties (String -> Object)
         * @param name property name
         * @param value property value
         */
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
