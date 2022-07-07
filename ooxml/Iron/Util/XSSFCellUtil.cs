using NPOI.SS.UserModel;
using NPOI.Util;
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
            ROTATION
        });
        private static readonly ISet<string> intValues = new HashSet<string>(new string[]{
            FONT
        });
        private static readonly ISet<string> booleanValues = new HashSet<string>(new string[]{
            LOCKED,
            HIDDEN,
            WRAP_TEXT,
            SHRINK_TO_FIT
        });
        private static readonly ISet<string> borderTypeValues = new HashSet<string>(new string[]{
            BORDER_BOTTOM,
            BORDER_LEFT,
            BORDER_RIGHT,
            BORDER_TOP
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

        public static void SetCellStyleProperties(ICell cell, XSSFWorkbook workbook, Dictionary<string, object> properties)
        {
            ICellStyle originalStyle = cell.CellStyle;
            ICellStyle newStyle = null;
            Dictionary<string, object> values = GetFormatProperties(originalStyle as XSSFCellStyle);

            PutAll(properties, values);

            int numberCellStyles = workbook.NumCellStyles;

            for (short i = 0; i < numberCellStyles; i++)
            {
                ICellStyle wbStyle = workbook.GetCellStyleAt(i);

                Dictionary<string, object> wbStyleMap = GetFormatProperties(wbStyle as XSSFCellStyle);

                if (values.Keys.Count != wbStyleMap.Keys.Count)
                {
                    continue;
                }

                bool stylesAreEqual = true;

                foreach (string key in values.Keys)
                {
                    if (!wbStyleMap.ContainsKey(key))
                    {
                        stylesAreEqual = false;
                        break;
                    }

                    if (values[key] == null && wbStyleMap[key] == null)
                    {
                        continue;
                    }

                    var wbVal = wbStyleMap[key];
                    var newVal = values[key];

                    if (newVal != null && newVal.Equals(wbVal))
                    {
                        continue;
                    }

                    stylesAreEqual = false;
                    break;
                }

                if (stylesAreEqual)
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

        /**
         * Copies the entries in src to dest, using the preferential data type
         * so that maps can be compared for equality
         *
         * @param src the property map to copy from (read-only)
         * @param dest the property map to copy into
         * @since POI 3.15 beta 3
         */
        private static void PutAll(Dictionary<string, object> src, Dictionary<string, object> dest)
        {
            foreach (string key in src.Keys)
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
                else
                {
                    //if (log.check(POILogger.INFO))
                    //{
                    //    log.log(POILogger.INFO, "Ignoring unrecognized CellUtil format properties key: " + key);
                    //}
                }

                /*
                 * BOTTOM_BORDER_COLOR,
                    LEFT_BORDER_COLOR,
                    RIGHT_BORDER_COLOR,
                    TOP_BORDER_COLOR,
                    FILL_FOREGROUND_COLOR,
                    FILL_BACKGROUND_COLOR,
                */
            }
        }

        private static Dictionary<string, object> GetFormatProperties(XSSFCellStyle style)
        {
            Dictionary<string, object> properties = new Dictionary<string, object>();
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
            PutShort(properties, FILL_PATTERN, (short)style.FillPattern);
            PutInt(properties, FONT, (int)style.FontIndex);
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

        /**
         * Utility method that puts the given value to the given map.
         *
         * @param properties map of properties (String -> Object)
         * @param name property name
         * @param value property value
         */
        private static void Put(Dictionary<String, Object> properties, String name, Object value)
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

            style.VerticalAlignment = GetVerticalAlignment(properties, VERTICAL_ALIGNMENT);
            style.WrapText = GetBoolean(properties, WRAP_TEXT);
        }

        /**
         * Utility method that returns the named short value form the given map.
         * 
         * @param properties map of named properties (String -> Object)
         * @param name property name
         * @return zero if the property does not exist, or is not a {@link Short}.
         */
        private static short GetShort(Dictionary<string, object> properties, string name)
        {
            object value = properties[name];
            return short.TryParse(value.ToString(), out short result) ? result : (short)0;
        }

        /**
         * Utility method that returns the named int value from the given map.
         *
         * @param properties map of named properties (String -> Object)
         * @param name property name
         * @return zero if the property does not exist, or is not a {@link Integer}
         *         otherwise the property value
         */
        private static int GetInt(Dictionary<string, object> properties, string name)
        {
            object value = properties[name];
            if (Number.IsNumber(value))
            {
                return int.Parse(value.ToString());
            }
            return 0;
        }

        /**
	     * Utility method that returns the named BorderStyle value form the given map.
	     *
	     * @param properties map of named properties (String -> Object)
	     * @param name property name
	     * @return Border style if set, otherwise {@link BorderStyle#NONE}
	     */
        private static BorderStyle GetBorderStyle(Dictionary<string, object> properties, string name)
        {
            object value = properties[name];
            BorderStyle border;
            if (value is BorderStyle borderStyle)
            {
                border = borderStyle;
            }
            else if (value is short || value is int)
            {
                short code = short.Parse(value.ToString());
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

        /**
         * Utility method that returns the named FillPattern value from the given map.
         *
         * @param properties map of named properties (String -> Object)
         * @param name property name
         * @return FillPattern style if set, otherwise {@link FillPattern#NO_FILL}
         * @since POI 3.15 beta 3
         */
        private static FillPattern GetFillPattern(Dictionary<string, object> properties, string name)
        {
            object value = properties[name];
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

        /**
         * Utility method that returns the named HorizontalAlignment value from the given map.
         *
         * @param properties map of named properties (String -> Object)
         * @param name property name
         * @return HorizontalAlignment style if set, otherwise {@link HorizontalAlignment#GENERAL}
         * @since POI 3.15 beta 3
         */
        private static HorizontalAlignment GetHorizontalAlignment(Dictionary<string, object> properties, string name)
        {
            object value = properties[name];
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

        /**
         * Utility method that returns the named VerticalAlignment value from the given map.
         *
         * @param properties map of named properties (String -> Object)
         * @param name property name
         * @return VerticalAlignment style if set, otherwise {@link VerticalAlignment#BOTTOM}
         * @since POI 3.15 beta 3
         */
        private static VerticalAlignment GetVerticalAlignment(Dictionary<string, object> properties, string name)
        {
            object value = properties[name];
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

        /**
         * Utility method that returns the named boolean value form the given map.
         *
         * @param properties map of properties (String -> Object)
         * @param name property name
         * @return false if the property does not exist, or is not a {@link Boolean}.
         */
        private static bool GetBoolean(Dictionary<string, object> properties, string name)
        {
            object value = properties[name];
            if (bool.TryParse(value.ToString(), out bool result))
                return result;

            return false;
        }
        private static string GetString(Dictionary<string, object> properties, string name)
        {
            object value = properties[name];
            return value?.ToString();
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

        /**
         * Utility method that puts the named short value to the given map.
         *
         * @param properties map of properties (String -> Object)
         * @param name property name
         * @param value property value
         */
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
