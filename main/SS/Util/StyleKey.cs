/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.Util
{
    using System;
    using System.Collections.Generic;
    using NPOI.SS.UserModel;

    /// <summary>
    /// Immutable value-type key capturing every comparable attribute of an
    /// <see cref="ICellStyle"/>.  Used as the dictionary key in
    /// <see cref="StyleCache"/> to provide O(1) style lookup.
    /// </summary>
    internal readonly struct StyleKey : IEquatable<StyleKey>
    {
        // ── shorts ───────────────────────────────────────────────────────────
        private readonly short _dataFormat;
        private readonly short _indention;
        private readonly short _rotation;
        private readonly short _leftBorderColor;
        private readonly short _rightBorderColor;
        private readonly short _topBorderColor;
        private readonly short _bottomBorderColor;
        private readonly short _fillForegroundColor;
        private readonly short _fillBackgroundColor;
        private readonly short _borderDiagonalColor;

        // ── int (font index stored as int in property map) ────────────────────
        private readonly int _fontIndex;

        // ── enums stored as int ──────────────────────────────────────────────
        private readonly int _alignment;
        private readonly int _verticalAlignment;
        private readonly int _fillPattern;
        private readonly int _borderLeft;
        private readonly int _borderRight;
        private readonly int _borderTop;
        private readonly int _borderBottom;
        private readonly int _borderDiagonalLineStyle;
        private readonly int _borderDiagonal;
        private readonly int _readingOrder;

        // ── bools packed into a single byte ──────────────────────────────────
        // bit 0 = IsHidden, bit 1 = IsLocked, bit 2 = WrapText,
        // bit 3 = ShrinkToFit, bit 4 = IsQuotePrefixed
        private readonly byte _flags;

        private StyleKey(
            short dataFormat, short indention, short rotation,
            short leftBorderColor, short rightBorderColor,
            short topBorderColor, short bottomBorderColor,
            short fillForegroundColor, short fillBackgroundColor,
            short borderDiagonalColor,
            int fontIndex,
            int alignment, int verticalAlignment, int fillPattern,
            int borderLeft, int borderRight, int borderTop, int borderBottom,
            int borderDiagonalLineStyle, int borderDiagonal, int readingOrder,
            byte flags)
        {
            _dataFormat = dataFormat;
            _indention = indention;
            _rotation = rotation;
            _leftBorderColor = leftBorderColor;
            _rightBorderColor = rightBorderColor;
            _topBorderColor = topBorderColor;
            _bottomBorderColor = bottomBorderColor;
            _fillForegroundColor = fillForegroundColor;
            _fillBackgroundColor = fillBackgroundColor;
            _borderDiagonalColor = borderDiagonalColor;
            _fontIndex = fontIndex;
            _alignment = alignment;
            _verticalAlignment = verticalAlignment;
            _fillPattern = fillPattern;
            _borderLeft = borderLeft;
            _borderRight = borderRight;
            _borderTop = borderTop;
            _borderBottom = borderBottom;
            _borderDiagonalLineStyle = borderDiagonalLineStyle;
            _borderDiagonal = borderDiagonal;
            _readingOrder = readingOrder;
            _flags = flags;
        }

        /// <summary>
        /// Creates a <see cref="StyleKey"/> by reading all properties directly
        /// from a live <see cref="ICellStyle"/> instance.
        /// </summary>
        public static StyleKey From(ICellStyle style)
        {
            byte flags = 0;
            if (style.IsHidden) flags |= 1;
            if (style.IsLocked) flags |= 2;
            if (style.WrapText) flags |= 4;
            if (style.ShrinkToFit) flags |= 8;
            if (style.IsQuotePrefixed) flags |= 16;

            return new StyleKey(
                style.DataFormat,
                style.Indention,
                style.Rotation,
                style.LeftBorderColor,
                style.RightBorderColor,
                style.TopBorderColor,
                style.BottomBorderColor,
                style.FillForegroundColor,
                style.FillBackgroundColor,
                style.BorderDiagonalColor,
                (int)style.FontIndex,
                (int)style.Alignment,
                (int)style.VerticalAlignment,
                (int)style.FillPattern,
                (int)style.BorderLeft,
                (int)style.BorderRight,
                (int)style.BorderTop,
                (int)style.BorderBottom,
                (int)style.BorderDiagonalLineStyle,
                (int)style.BorderDiagonal,
                (int)style.ReadingOrder,
                flags);
        }

        /// <summary>
        /// Creates a <see cref="StyleKey"/> directly from the property map
        /// produced by <c>GetFormatProperties</c> + <c>PutAll</c>, avoiding
        /// any temporary <see cref="ICellStyle"/> allocation on the hot path.
        /// Properties not tracked in the map (e.g. BorderDiagonal, ReadingOrder)
        /// default to zero/false, matching the values set by
        /// <c>SetFormatProperties</c> on a freshly created style.
        /// </summary>
        public static StyleKey FromPropertyMap(Dictionary<string, object> values)
        {
            byte flags = 0;
            if (CellUtil.GetBoolean(values, CellUtil.HIDDEN)) flags |= 1;
            if (CellUtil.GetBoolean(values, CellUtil.LOCKED)) flags |= 2;
            if (CellUtil.GetBoolean(values, CellUtil.WRAP_TEXT)) flags |= 4;
            // ShrinkToFit and IsQuotePrefixed are not currently in the property
            // map (commented out in GetFormatProperties); default to 0.

            return new StyleKey(
                CellUtil.GetShort(values, CellUtil.DATA_FORMAT),
                CellUtil.GetShort(values, CellUtil.INDENTION),
                CellUtil.GetShort(values, CellUtil.ROTATION),
                CellUtil.GetShort(values, CellUtil.LEFT_BORDER_COLOR),
                CellUtil.GetShort(values, CellUtil.RIGHT_BORDER_COLOR),
                CellUtil.GetShort(values, CellUtil.TOP_BORDER_COLOR),
                CellUtil.GetShort(values, CellUtil.BOTTOM_BORDER_COLOR),
                CellUtil.GetShort(values, CellUtil.FILL_FOREGROUND_COLOR),
                CellUtil.GetShort(values, CellUtil.FILL_BACKGROUND_COLOR),
                0, // BorderDiagonalColor not tracked in property map
                CellUtil.GetInt(values, CellUtil.FONT),
                (int)CellUtil.GetHorizontalAlignment(values, CellUtil.ALIGNMENT),
                (int)CellUtil.GetVerticalAlignment(values, CellUtil.VERTICAL_ALIGNMENT),
                (int)CellUtil.GetFillPattern(values, CellUtil.FILL_PATTERN),
                (int)CellUtil.GetBorderStyle(values, CellUtil.BORDER_LEFT),
                (int)CellUtil.GetBorderStyle(values, CellUtil.BORDER_RIGHT),
                (int)CellUtil.GetBorderStyle(values, CellUtil.BORDER_TOP),
                (int)CellUtil.GetBorderStyle(values, CellUtil.BORDER_BOTTOM),
                0, // BorderDiagonalLineStyle not tracked in property map
                0, // BorderDiagonal not tracked in property map
                0, // ReadingOrder not tracked in property map
                flags);
        }

        /// <inheritdoc/>
        public bool Equals(StyleKey other)
        {
            return _dataFormat == other._dataFormat
                && _indention == other._indention
                && _rotation == other._rotation
                && _leftBorderColor == other._leftBorderColor
                && _rightBorderColor == other._rightBorderColor
                && _topBorderColor == other._topBorderColor
                && _bottomBorderColor == other._bottomBorderColor
                && _fillForegroundColor == other._fillForegroundColor
                && _fillBackgroundColor == other._fillBackgroundColor
                && _borderDiagonalColor == other._borderDiagonalColor
                && _fontIndex == other._fontIndex
                && _alignment == other._alignment
                && _verticalAlignment == other._verticalAlignment
                && _fillPattern == other._fillPattern
                && _borderLeft == other._borderLeft
                && _borderRight == other._borderRight
                && _borderTop == other._borderTop
                && _borderBottom == other._borderBottom
                && _borderDiagonalLineStyle == other._borderDiagonalLineStyle
                && _borderDiagonal == other._borderDiagonal
                && _readingOrder == other._readingOrder
                && _flags == other._flags;
        }

        /// <inheritdoc/>
        public override bool Equals(object obj) => obj is StyleKey key && Equals(key);

        public static bool operator ==(StyleKey left, StyleKey right) => left.Equals(right);
        public static bool operator !=(StyleKey left, StyleKey right) => !left.Equals(right);

        /// <inheritdoc/>
        public override int GetHashCode()
        {
            unchecked
            {
                int h = 17;
                h = h * 31 + _dataFormat;
                h = h * 31 + _indention;
                h = h * 31 + _rotation;
                h = h * 31 + _leftBorderColor;
                h = h * 31 + _rightBorderColor;
                h = h * 31 + _topBorderColor;
                h = h * 31 + _bottomBorderColor;
                h = h * 31 + _fillForegroundColor;
                h = h * 31 + _fillBackgroundColor;
                h = h * 31 + _borderDiagonalColor;
                h = h * 31 + _fontIndex;
                h = h * 31 + _alignment;
                h = h * 31 + _verticalAlignment;
                h = h * 31 + _fillPattern;
                h = h * 31 + _borderLeft;
                h = h * 31 + _borderRight;
                h = h * 31 + _borderTop;
                h = h * 31 + _borderBottom;
                h = h * 31 + _borderDiagonalLineStyle;
                h = h * 31 + _borderDiagonal;
                h = h * 31 + _readingOrder;
                h = h * 31 + _flags;
                return h;
            }
        }
    }
}
