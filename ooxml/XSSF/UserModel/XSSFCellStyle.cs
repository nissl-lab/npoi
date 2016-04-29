/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using System.Xml;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel.Extensions;
using NPOI.XSSF.Model;
namespace NPOI.XSSF.UserModel
{


    /**
     *
     * High level representation of the the possible formatting information for the contents of the cells on a sheet in a
     * SpreadsheetML document.
     *
     * @see NPOI.xssf.usermodel.XSSFWorkbook#CreateCellStyle()
     * @see NPOI.xssf.usermodel.XSSFWorkbook#getCellStyleAt(short)
     * @see NPOI.xssf.usermodel.XSSFCell#setCellStyle(NPOI.ss.usermodel.CellStyle)
     */
    public class XSSFCellStyle : ICellStyle
    {

        private int _cellXfId;
        private StylesTable _stylesSource;
        private CT_Xf _cellXf;
        private CT_Xf _cellStyleXf;
        private XSSFFont _font;
        private XSSFCellAlignment _cellAlignment;
        private ThemesTable _theme;

        /**
         * Creates a Cell Style from the supplied parts
         * @param cellXfId The main XF for the cell. Must be a valid 0-based index into the XF table
         * @param cellStyleXfId Optional, style xf. A value of <code>-1</code> means no xf.
         * @param stylesSource Styles Source to work off
         */
        public XSSFCellStyle(int cellXfId, int cellStyleXfId, StylesTable stylesSource, ThemesTable theme)
        {
            _cellXfId = cellXfId;
            _stylesSource = stylesSource;
            _cellXf = stylesSource.GetCellXfAt(this._cellXfId);
            _cellStyleXf = cellStyleXfId == -1 ? null : stylesSource.GetCellStyleXfAt(cellStyleXfId);
            _theme = theme;
        }

        /**
         * Used so that StylesSource can figure out our location
         */

        public CT_Xf GetCoreXf()
        {
            return _cellXf;
        }

        /**
         * Used so that StylesSource can figure out our location
         */

        public CT_Xf GetStyleXf()
        {
            return _cellStyleXf;
        }

        /// <summary>
        /// Creates an empty Cell Style
        /// </summary>
        /// <param name="stylesSource"></param>
        public XSSFCellStyle(StylesTable stylesSource)
        {
            _stylesSource = stylesSource;
            // We need a new CT_Xf for the main styles
            // TODO decide on a style ctxf
            _cellXf = new CT_Xf();
            _cellStyleXf = null;
        }

        /**
         * Verifies that this style belongs to the supplied Workbook
         *  Styles Source.
         * Will throw an exception if it belongs to a different one.
         * This is normally called when trying to assign a style to a
         *  cell, to ensure the cell and the style are from the same
         *  workbook (if they're not, it won't work)
         * @throws ArgumentException if there's a workbook mis-match
         */
        public void VerifyBelongsToStylesSource(StylesTable src)
        {
            if (this._stylesSource != src)
            {
                throw new ArgumentException("This Style does not belong to the supplied Workbook Stlyes Source. Are you trying to assign a style from one workbook to the cell of a differnt workbook?");
            }
        }

        /**
         * Clones all the style information from another
         *  XSSFCellStyle, onto this one. This
         *  XSSFCellStyle will then have all the same
         *  properties as the source, but the two may
         *  be edited independently.
         * Any stylings on this XSSFCellStyle will be lost!
         *
         * The source XSSFCellStyle could be from another
         *  XSSFWorkbook if you like. This allows you to
         *  copy styles from one XSSFWorkbook to another.
         */
        public void CloneStyleFrom(ICellStyle source)
        {
            if (source is XSSFCellStyle)
            {
                XSSFCellStyle src = (XSSFCellStyle)source;

                // Is it on our Workbook?
                if (src._stylesSource == _stylesSource)
                {
                    // Nice and easy
                    _cellXf = src.GetCoreXf().Copy();
                    _cellStyleXf = src.GetStyleXf().Copy();
                }
                else
                {
                    // Copy the style
                    try
                    {
                        // Remove any children off the current style, to
                        //  avoid orphaned nodes
                        if (_cellXf.IsSetAlignment())
                            _cellXf.UnsetAlignment();
                        if (_cellXf.IsSetExtLst())
                            _cellXf.UnsetExtLst();

                        // Create a new Xf with the same contents
                        _cellXf =
                              src.GetCoreXf().Copy();

                        // bug 56295: ensure that the fills is available and set correctly
                        CT_Fill fill = CT_Fill.Parse(src.GetCTFill().ToString());
                        AddFill(fill);

                        // Swap it over
                        _stylesSource.ReplaceCellXfAt(_cellXfId, _cellXf);
                    }
                    catch (XmlException e)
                    {
                        throw new POIXMLException(e);
                    }

                    // Copy the format
                    String fmt = src.GetDataFormatString();
                    DataFormat = (
                          (new XSSFDataFormat(_stylesSource)).GetFormat(fmt)
                    );

                    // Copy the font
                    try
                    {
                        CT_Font ctFont =
                              src.GetFont().GetCTFont().Clone();
                        XSSFFont font = new XSSFFont(ctFont);
                        font.RegisterTo(_stylesSource);
                        SetFont(font);
                    }
                    catch (XmlException e)
                    {
                        throw new POIXMLException(e);
                    }
                }

                // Clear out cached details
                _font = null;
                _cellAlignment = null;
            }
            else
            {
                throw new ArgumentException("Can only clone from one XSSFCellStyle to another, not between HSSFCellStyle and XSSFCellStyle");
            }
        }
        private void AddFill(CT_Fill fill)
        {
            int idx = _stylesSource.PutFill(new XSSFCellFill(fill));

            _cellXf.fillId = (uint)(idx);
            _cellXf.applyFill = (true);
        }
        public HorizontalAlignment Alignment
        {
            get
            {
                return GetAlignmentEnum();
            }
            set
            {
                GetCellAlignment().Horizontal = value;
            }
        }

        /// <summary>
        /// Get the type of horizontal alignment for the cell
        /// </summary>
        /// <returns>the type of alignment</returns>
        internal HorizontalAlignment GetAlignmentEnum()
        {
            CT_CellAlignment align = _cellXf.alignment;
            if (align != null && align.IsSetHorizontal())
            {
                return (HorizontalAlignment)align.horizontal;
            }
            return HorizontalAlignment.General;
        }

        public BorderStyle BorderBottom
        {
            get
            {
                if (!_cellXf.applyBorder) return BorderStyle.None;

                int idx = (int)_cellXf.borderId;
                CT_Border ct = _stylesSource.GetBorderAt(idx).GetCTBorder();
                if (!ct.IsSetBottom())
                {
                    return BorderStyle.None;
                }
                else
                {
                    return (BorderStyle)ct.bottom.style;
                }
            }
            set
            {
                CT_Border ct = GetCTBorder();
                CT_BorderPr pr = ct.IsSetBottom() ? ct.bottom : ct.AddNewBottom();
                if (value == BorderStyle.None) 
                    ct.UnsetBottom();
                else 
                    pr.style = (ST_BorderStyle)value;

                int idx = _stylesSource.PutBorder(new XSSFCellBorder(ct, _theme));

                _cellXf.borderId = (uint)idx;
                _cellXf.applyBorder = (true);
            }
        }


        public BorderStyle BorderLeft
        {
            get
            {
                if (!_cellXf.applyBorder) return BorderStyle.None;

                int idx = (int)_cellXf.borderId;
                CT_Border ct = _stylesSource.GetBorderAt(idx).GetCTBorder();
                if (!ct.IsSetLeft())
                {
                    return BorderStyle.None;
                }
                else
                {
                    return (BorderStyle)ct.left.style;
                }
            }
            set
            {
                CT_Border ct = GetCTBorder();
                CT_BorderPr pr = ct.IsSetLeft() ? ct.left : ct.AddNewLeft();
                if (value == BorderStyle.None) ct.unsetLeft();
                else pr.style = (ST_BorderStyle)value;

                int idx = _stylesSource.PutBorder(new XSSFCellBorder(ct, _theme));

                _cellXf.borderId = (uint)idx;
                _cellXf.applyBorder = (true);

            }
        }

        /// <summary>
        /// Get the type of border to use for the right border of the cell
        /// </summary>
        public BorderStyle BorderRight
        {
            get
            {
                if (!_cellXf.applyBorder) return BorderStyle.None;

                int idx = (int)_cellXf.borderId;
                CT_Border ct = _stylesSource.GetBorderAt(idx).GetCTBorder();
                if (!ct.IsSetRight())
                {
                    return BorderStyle.None;
                }
                else
                {
                    return (BorderStyle)ct.right.style;
                }
            }
            set
            {
                CT_Border ct = GetCTBorder();
                CT_BorderPr pr = ct.IsSetRight() ? ct.right : ct.AddNewRight();
                if (value == BorderStyle.None) ct.unsetRight();
                else pr.style = (ST_BorderStyle)value;

                int idx = _stylesSource.PutBorder(new XSSFCellBorder(ct, _theme));

                _cellXf.borderId = (uint)idx;
                _cellXf.applyBorder = (true);
            }
        }

        public BorderStyle BorderTop
        {
            get
            {
                if (!_cellXf.applyBorder) return BorderStyle.None;

                int idx = (int)_cellXf.borderId;
                CT_Border ct = _stylesSource.GetBorderAt(idx).GetCTBorder();
                if (!ct.IsSetTop())
                {
                    return BorderStyle.None;
                }
                else
                {
                    return (BorderStyle)ct.top.style;
                }
            }
            set
            {
                CT_Border ct = GetCTBorder();
                CT_BorderPr pr = ct.IsSetTop() ? ct.top : ct.AddNewTop();
                if (value == BorderStyle.None) ct.unsetTop();
                else pr.style = (ST_BorderStyle)value;

                int idx = _stylesSource.PutBorder(new XSSFCellBorder(ct, _theme));

                _cellXf.borderId = (uint)idx;
                _cellXf.applyBorder = (true);
            }
        }

        /**
         * Get the color to use for the bottom border
         * Color is optional. When missing, IndexedColors.Automatic is implied.
         * @return the index of the color defInition, default value is {@link NPOI.ss.usermodel.IndexedColors#AUTOMATIC}
         * @see NPOI.ss.usermodel.IndexedColors
         */
        public short BottomBorderColor
        {
            get
            {
                XSSFColor clr = BottomBorderXSSFColor;
                return clr == null ? IndexedColors.Black.Index : clr.Indexed;
            }
            set
            {
                XSSFColor clr = new XSSFColor();
                clr.Indexed = value;
                SetBottomBorderColor(clr);
            }
        }

        /**
         * Get the color to use for the bottom border as a {@link XSSFColor}
         *
         * @return the used color or <code>null</code> if not Set
         */
        public XSSFColor BottomBorderXSSFColor
        {
            get
            {
                if (!_cellXf.applyBorder) return null;

                int idx = (int)_cellXf.borderId;
                XSSFCellBorder border = _stylesSource.GetBorderAt(idx);

                return border.GetBorderColor(BorderSide.BOTTOM);
            }
        }

        /**
         * Get the index of the number format (numFmt) record used by this cell format.
         *
         * @return the index of the number format
         */
        public short DataFormat
        {
            get
            {
                return (short)_cellXf.numFmtId;
            }
            set
            {
                // XSSF supports >32,767 formats
                SetDataFormat(value & 0xFFFF);
            }
        }
        /**
         * Set the index of a data format
         *
         * @param fmt the index of a data format
         */
            public void SetDataFormat(int fmt)
            {
                _cellXf.applyNumberFormat = (true);
                _cellXf.numFmtId = (uint)(fmt);
            }
        /**
         * Get the contents of the format string, by looking up
         * the StylesSource
         *
         * @return the number format string
         */
        public String GetDataFormatString()
        {
            int idx = DataFormat;
            return new XSSFDataFormat(_stylesSource).GetFormat((short)idx);
        }

        /// <summary>
        /// Get the background fill color.
        /// Note - many cells are actually filled with a foreground fill, not a background fill
        /// </summary>
        public short FillBackgroundColor
        {
            get
            {
                XSSFColor clr = (XSSFColor)this.FillBackgroundColorColor;
                return clr == null ? IndexedColors.Automatic.Index : clr.Indexed;
            }
            set
            {
                XSSFColor clr = new XSSFColor();
                clr.Indexed = (value);
                SetFillBackgroundColor(clr);
            }
        }

        /**
         * Get the background fill color.
         * <p>
         * Note - many cells are actually Filled with a foreground
         *  Fill, not a background fill - see {@link #getFillForegroundColor()}
         * </p>
         * @see NPOI.xssf.usermodel.XSSFColor#getRgb()
         * @return XSSFColor - fill color or <code>null</code> if not Set
         */
        public IColor FillBackgroundColorColor
        {
            get
            {
                return this.FillBackgroundXSSFColor;
            }
            set
            {
                this.FillBackgroundXSSFColor = (XSSFColor)value;
            }
        }
        public XSSFColor FillBackgroundXSSFColor
        {
            get
            {
                // bug 56295: handle missing applyFill attribute as "true" because Excel does as well
                if (_cellXf.IsSetApplyFill() && !_cellXf.applyFill) return null;

                int fillIndex = (int)_cellXf.fillId;
                XSSFCellFill fg = _stylesSource.GetFillAt(fillIndex);

                XSSFColor fillBackgroundColor = fg.GetFillBackgroundColor();
                if (fillBackgroundColor != null && _theme != null)
                {
                    _theme.InheritFromThemeAsRequired(fillBackgroundColor);
                }
                return fillBackgroundColor;
            }
            set
            {
                CT_Fill ct = GetCTFill();
                CT_PatternFill ptrn = ct.patternFill;
                if (value == null)
                {
                    if (ptrn != null) ptrn.UnsetBgColor();
                }
                else
                {
                    if (ptrn == null) ptrn = ct.AddNewPatternFill();
                    ptrn.bgColor = (value.GetCTColor());
                }

                AddFill(ct);
            }
        }
        /**
         * Get the foreground fill color.
         * <p>
         * Many cells are Filled with this, instead of a
         *  background color ({@link #getFillBackgroundColor()})
         * </p>
         * @see IndexedColors
         * @return fill color, default value is {@link NPOI.ss.usermodel.IndexedColors#AUTOMATIC}
         */
        public short FillForegroundColor
        {
            get
            {
                XSSFColor clr = (XSSFColor)this.FillForegroundColorColor;
                return clr == null ? IndexedColors.Automatic.Index : clr.Indexed;
            }
            set
            {
                XSSFColor clr = new XSSFColor();
                clr.Indexed = (value);
                SetFillForegroundColor(clr);
            }
        }

        /// <summary>
        /// Get the foreground fill color.
        /// </summary>
        public IColor FillForegroundColorColor
        {
            get
            {
                return this.FillForegroundXSSFColor;
            }
            set
            {
                this.FillForegroundXSSFColor = (XSSFColor)value;
            }
        }

        /// <summary>
        /// Get the foreground fill color.
        /// </summary>
        public XSSFColor FillForegroundXSSFColor
        {
            get
            {
                // bug 56295: handle missing applyFill attribute as "true" because Excel does as well
                if (_cellXf.IsSetApplyFill() && !_cellXf.applyFill) return null;

                int fillIndex = (int)_cellXf.fillId;
                XSSFCellFill fg = _stylesSource.GetFillAt(fillIndex);

                XSSFColor fillForegroundColor = fg.GetFillForegroundColor();
                if (fillForegroundColor != null && _theme != null)
                {
                    _theme.InheritFromThemeAsRequired(fillForegroundColor);
                }
                return fillForegroundColor;
            }
            set
            {
                CT_Fill ct = GetCTFill();

                CT_PatternFill ptrn = ct.patternFill;
                if (value == null)
                {
                    if (ptrn != null) ptrn.UnsetFgColor();
                }
                else
                {
                    if (ptrn == null) ptrn = ct.AddNewPatternFill();
                    ptrn.fgColor = (value.GetCTColor());
                }

                AddFill(ct);
            }
        }
        public FillPattern FillPattern
        {
            get
            {
                // bug 56295: handle missing applyFill attribute as "true" because Excel does as well
                if (_cellXf.IsSetApplyFill() && !_cellXf.applyFill) return 0;

                int FillIndex = (int)_cellXf.fillId;
                XSSFCellFill fill = _stylesSource.GetFillAt(FillIndex);

                ST_PatternType ptrn = fill.GetPatternType();
                if(ptrn == ST_PatternType.none) return FillPattern.NoFill;

                return (FillPattern)((int)ptrn);
            }
            set
            {
                CT_Fill ct = GetCTFill();
                CT_PatternFill ptrn = ct.IsSetPatternFill() ? ct.GetPatternFill() : ct.AddNewPatternFill();
                if (value == FillPattern.NoFill && ptrn.IsSetPatternType())
                    ptrn.UnsetPatternType();
                else ptrn.patternType = (ST_PatternType)(value);

                AddFill(ct);
            }
        }


        /**
        * Gets the font for this style
        * @return Font - font
        */
        public XSSFFont GetFont()
        {
            if (_font == null)
            {
                _font = _stylesSource.GetFontAt(FontId);
            }
            return _font;
        }

        /**
         * Gets the index of the font for this style
         *
         * @return short - font index
         * @see NPOI.xssf.usermodel.XSSFWorkbook#getFontAt(short)
         */
        public short FontIndex
        {
            get
            {
                return (short)FontId;
            }
        }

        /**
         * Get whether the cell's using this style are to be hidden
         *
         * @return bool -  whether the cell using this style is hidden
         */
        public bool IsHidden
        {
            get
            {
                if (!_cellXf.IsSetProtection() || !_cellXf.protection.IsSetHidden())
                {
                    return false;
                }
                return _cellXf.protection.hidden;
            }
            set
            {
                if (!_cellXf.IsSetProtection())
                {
                    _cellXf.AddNewProtection();
                }
                _cellXf.protection.hidden = (value);
            }
        }

        /**
         * Get the number of spaces to indent the text in the cell
         *
         * @return indent - number of spaces
         */
        public short Indention
        {
            get
            {
                CT_CellAlignment align = _cellXf.alignment;
                return (short)(align == null ? 0 : align.indent);
            }
            set
            {
                GetCellAlignment().Indent = value;
            }
        }

        /**
         * Get the index within the StylesTable (sequence within the collection of CT_Xf elements)
         *
         * @return unique index number of the underlying record this style represents
         */
        public short Index
        {
            get
            {
                return (short)this._cellXfId;
            }
        }
        protected internal int UIndex
        {
            get
            {
                return this._cellXfId;
            }
        }
        /**
         * Get the color to use for the left border
         *
         * @return the index of the color defInition, default value is {@link NPOI.ss.usermodel.IndexedColors#BLACK}
         * @see NPOI.ss.usermodel.IndexedColors
         */
        public short LeftBorderColor
        {
            get
            {
                XSSFColor clr = LeftBorderXSSFColor;
                return clr == null ? IndexedColors.Black.Index : clr.Indexed;
            }
            set
            {
                XSSFColor clr = new XSSFColor();
                clr.Indexed = (value);
                SetLeftBorderColor(clr);
            }
        }
        public XSSFColor DiagonalBorderXSSFColor
        {
            get
            {
                if (!_cellXf.applyBorder) return null;

                int idx = (int)_cellXf.borderId;
                XSSFCellBorder border = _stylesSource.GetBorderAt(idx);

                return border.GetBorderColor(BorderSide.DIAGONAL);
            }
        }
        /**
         * Get the color to use for the left border
         *
         * @return the index of the color defInition or <code>null</code> if not Set
         * @see NPOI.ss.usermodel.IndexedColors
         */
        public XSSFColor LeftBorderXSSFColor
        {
            get
            {
                if (!_cellXf.applyBorder) return null;

                int idx = (int)_cellXf.borderId;
                XSSFCellBorder border = _stylesSource.GetBorderAt(idx);

                return border.GetBorderColor(BorderSide.LEFT);
            }
        }
 
        /// <summary>
        /// Get whether the cell's using this style are locked
        /// </summary>
        public bool IsLocked
        {
            get
            {

                if (!_cellXf.IsSetProtection())
                {
                    return true;
                }
                return _cellXf.protection.locked;
            }
            set
            {
                if (!_cellXf.IsSetProtection())
                {
                    _cellXf.AddNewProtection();
                }
                _cellXf.protection.locked = value;
            }
        }

        /// <summary>
        /// Get the color to use for the right border
        /// </summary>
        public short RightBorderColor
        {
            get
            {
                XSSFColor clr = RightBorderXSSFColor;
                return clr == null ? IndexedColors.Black.Index : clr.Indexed;
            }
            set
            {
                XSSFColor clr = new XSSFColor();
                clr.Indexed = (value);
                SetRightBorderColor(clr);
            }
        }
        /// <summary>
        /// Get the color to use for the right border
        /// </summary>
        /// <returns></returns>
        public XSSFColor RightBorderXSSFColor
        {
            get
            {
                if (!_cellXf.applyBorder) return null;

                int idx = (int)_cellXf.borderId;
                XSSFCellBorder border = _stylesSource.GetBorderAt(idx);

                return border.GetBorderColor(BorderSide.RIGHT);
            }
        }
        /// <summary>
        /// Get the degree of rotation (between 0 and 180 degrees) for the text in the cell
        /// </summary>
        /// <example>
        /// Expressed in degrees. Values range from 0 to 180. The first letter of
        /// the text is considered the center-point of the arc.
        /// For 0 - 90, the value represents degrees above horizon. For 91-180 the degrees below the horizon is calculated as:
        /// <code>[degrees below horizon] = 90 - textRotation.</code>
        /// </example>
        public short Rotation
        {
            get
            {
                CT_CellAlignment align = _cellXf.alignment;
                return (short)(align == null ? 0 : align.textRotation);
            }
            set
            {
                GetCellAlignment().TextRotation = value;
            }
        }

        /**
         * Get the color to use for the top border
         *
         * @return the index of the color defInition, default value is {@link NPOI.ss.usermodel.IndexedColors#BLACK}
         * @see NPOI.ss.usermodel.IndexedColors
         */
        public short TopBorderColor
        {
            get
            {
                XSSFColor clr = TopBorderXSSFColor;
                return clr == null ? IndexedColors.Black.Index : clr.Indexed;
            }
            set
            {
                XSSFColor clr = new XSSFColor();
                clr.Indexed = (value);
                SetTopBorderColor(clr);
            }
        }
        /// <summary>
        /// Get the color to use for the top border
        /// </summary>
        /// <returns></returns>
        public XSSFColor TopBorderXSSFColor
        {
            get
            {
                if (!_cellXf.applyBorder) return null;

                int idx = (int)_cellXf.borderId;
                XSSFCellBorder border = _stylesSource.GetBorderAt(idx);

                return border.GetBorderColor(BorderSide.TOP);
            }
        }

        /// <summary>
        /// Get the type of vertical alignment for the cell
        /// </summary>
        public VerticalAlignment VerticalAlignment
        {
            get
            {
                return GetVerticalAlignmentEnum();
            }
            set
            {
                GetCellAlignment().Vertical = value;
            }
        }
        /// <summary>
        /// Get the type of vertical alignment for the cell
        /// </summary>
        /// <returns></returns>
        internal VerticalAlignment GetVerticalAlignmentEnum()
        {
            CT_CellAlignment align = _cellXf.alignment;
            if (align != null && align.IsSetVertical())
            {
                return (VerticalAlignment)align.vertical;
            }
            return VerticalAlignment.Bottom;
        }
        /// <summary>
        /// Whether the text in a cell should be line-wrapped within the cell.
        /// </summary>
        public bool WrapText
        {
            get
            {
                CT_CellAlignment align = _cellXf.alignment;
                return align != null && align.wrapText;
            }
            set
            {
                GetCellAlignment().WrapText = value;
            }
        }


        /**
         * Set the color to use for the bottom border
         *
         * @param color the color to use, null means no color
         */
        public void SetBottomBorderColor(XSSFColor color)
        {
            CT_Border ct = GetCTBorder();
            if (color == null && !ct.IsSetBottom()) return;

            CT_BorderPr pr = ct.IsSetBottom() ? ct.bottom : ct.AddNewBottom();
            if (color != null) pr.SetColor(color.GetCTColor());
            else pr.UnsetColor();

            int idx = _stylesSource.PutBorder(new XSSFCellBorder(ct, _theme));

            _cellXf.borderId = (uint)idx;
            _cellXf.applyBorder = (true);
        }
        /**
         * Set the background fill color represented as a {@link XSSFColor} value.
         * <p>
         * For example:
         * <pre>
         * cs.SetFillPattern(XSSFCellStyle.FINE_DOTS );
         * cs.SetFillBackgroundXSSFColor(new XSSFColor(java.awt.Color.RED));
         * </pre>
         * optionally a Foreground and background fill can be applied:
         * <i>Note: Ensure Foreground color is set prior to background</i>
         * <pre>
         * cs.SetFillPattern(XSSFCellStyle.FINE_DOTS );
         * cs.SetFillForegroundColor(new XSSFColor(java.awt.Color.BLUE));
         * cs.SetFillBackgroundColor(new XSSFColor(java.awt.Color.GREEN));
         * </pre>
         * or, for the special case of SOLID_FILL:
         * <pre>
         * cs.SetFillPattern(XSSFCellStyle.SOLID_FOREGROUND );
         * cs.SetFillForegroundColor(new XSSFColor(java.awt.Color.GREEN));
         * </pre>
         * It is necessary to set the fill style in order
         * for the color to be shown in the cell.
         *
         * @param color - the color to use
         */
        public void SetFillBackgroundColor(XSSFColor color)
        {
            CT_Fill ct = GetCTFill();
            CT_PatternFill ptrn = ct.GetPatternFill();
            if (color == null)
            {
                if (ptrn != null) ptrn.UnsetBgColor();
            }
            else
            {

                if (ptrn == null) ptrn = ct.AddNewPatternFill();
                ptrn.bgColor = color.GetCTColor();
            }

            int idx = _stylesSource.PutFill(new XSSFCellFill(ct));

            _cellXf.fillId = (uint)idx;
            _cellXf.applyFill = (true);
        }

        /**
        * Set the foreground fill color represented as a {@link XSSFColor} value.
         * <br/>
        * <i>Note: Ensure Foreground color is Set prior to background color.</i>
        * @param color the color to use
        * @see #setFillBackgroundColor(NPOI.xssf.usermodel.XSSFColor) )
        */
        public void SetFillForegroundColor(XSSFColor color)
        {
            CT_Fill ct = GetCTFill();

            CT_PatternFill ptrn = ct.GetPatternFill();
            if (color == null)
            {
                if (ptrn != null) ptrn.UnsetFgColor();
            }
            else
            {
                if (ptrn == null) ptrn = ct.AddNewPatternFill();
                ptrn.fgColor = (color.GetCTColor());
            }

            int idx = _stylesSource.PutFill(new XSSFCellFill(ct));

            _cellXf.fillId = (uint)idx;
            _cellXf.applyFill = (true);
        }

        /**
         * Get a <b>copy</b> of the currently used CT_Fill, if none is used, return a new instance.
         */
        public CT_Fill GetCTFill()
        {
            CT_Fill ct;
            // bug 56295: handle missing applyFill attribute as "true" because Excel does as well
            if (!_cellXf.IsSetApplyFill() || _cellXf.applyFill)
            {
                int FillIndex = (int)_cellXf.fillId;
                XSSFCellFill cf = _stylesSource.GetFillAt(FillIndex);

                ct = (CT_Fill)cf.GetCTFill().Copy();
            }
            else
            {
                ct = new CT_Fill();
            }
            return ct;
        }

        /**
         * Get a <b>copy</b> of the currently used CT_Border, if none is used, return a new instance.
         */
        public CT_Border GetCTBorder()
        {
            CT_Border ctBorder;
            if (_cellXf.applyBorder)
            {
                int idx = (int)_cellXf.borderId;
                XSSFCellBorder cf = _stylesSource.GetBorderAt(idx);

                ctBorder = (CT_Border)cf.GetCTBorder();
            }
            else
            {
                ctBorder = new CT_Border();
                ctBorder.AddNewLeft();
                ctBorder.AddNewRight();
                ctBorder.AddNewTop();
                ctBorder.AddNewBottom();
                ctBorder.AddNewDiagonal();
            }
            return ctBorder;
        }


        /**
         * Set the font for this style
         *
         * @param font  a font object Created or retreived from the XSSFWorkbook object
         * @see NPOI.xssf.usermodel.XSSFWorkbook#CreateFont()
         * @see NPOI.xssf.usermodel.XSSFWorkbook#getFontAt(short)
         */
        public void SetFont(IFont font)
        {
            if (font != null)
            {
                long index = font.Index;
                this._cellXf.fontId = (uint)index;
                this._cellXf.fontIdSpecified = true;
                this._cellXf.applyFont = (true);
            }
            else
            {
                this._cellXf.applyFont = (false);
            }
        }
        public void SetDiagonalBorderColor(XSSFColor color)
        {
            CT_Border ct = GetCTBorder();
            if (color == null && !ct.IsSetDiagonal()) return;

            CT_BorderPr pr = ct.IsSetDiagonal() ? ct.diagonal : ct.AddNewDiagonal();
            if (color != null) pr.color = (color.GetCTColor());
            else pr.UnsetColor();

            int idx = _stylesSource.PutBorder(new XSSFCellBorder(ct, _theme));

            _cellXf.borderId = (uint)idx;
            _cellXf.applyBorder = (true);
        }
        /**
         * Set the color to use for the left border as a {@link XSSFColor} value
         *
         * @param color the color to use
         */
        public void SetLeftBorderColor(XSSFColor color)
        {
            CT_Border ct = GetCTBorder();
            if (color == null && !ct.IsSetLeft()) return;

            CT_BorderPr pr = ct.IsSetLeft() ? ct.left : ct.AddNewLeft();
            if (color != null) pr.color = (color.GetCTColor());
            else pr.UnsetColor();

            int idx = _stylesSource.PutBorder(new XSSFCellBorder(ct, _theme));

            _cellXf.borderId = (uint)idx;
            _cellXf.applyBorder = (true);
        }

        /**
         * Set the color to use for the right border as a {@link XSSFColor} value
         *
         * @param color the color to use
         */
        public void SetRightBorderColor(XSSFColor color)
        {
            CT_Border ct = GetCTBorder();
            if (color == null && !ct.IsSetRight()) return;

            CT_BorderPr pr = ct.IsSetRight() ? ct.right : ct.AddNewRight();
            if (color != null) pr.color = (color.GetCTColor());
            else pr.UnsetColor();

            int idx = _stylesSource.PutBorder(new XSSFCellBorder(ct, _theme));

            _cellXf.borderId = (uint)(idx);
            _cellXf.applyBorder = (true);
        }




        /**
         * Set the color to use for the top border as a {@link XSSFColor} value
         *
         * @param color the color to use
         */
        public void SetTopBorderColor(XSSFColor color)
        {
            CT_Border ct = GetCTBorder();
            if (color == null && !ct.IsSetTop()) return;

            CT_BorderPr pr = ct.IsSetTop() ? ct.top : ct.AddNewTop();
            if (color != null) pr.color = color.GetCTColor();
            else pr.UnsetColor();

            int idx = _stylesSource.PutBorder(new XSSFCellBorder(ct, _theme));

            _cellXf.borderId = (uint)idx;
            _cellXf.applyBorder = (true);
        }

        /**
         * Set the type of vertical alignment for the cell
         *
         * @param align - align the type of alignment
         * @see NPOI.ss.usermodel.CellStyle#VERTICAL_TOP
         * @see NPOI.ss.usermodel.CellStyle#VERTICAL_CENTER
         * @see NPOI.ss.usermodel.CellStyle#VERTICAL_BOTTOM
         * @see NPOI.ss.usermodel.CellStyle#VERTICAL_JUSTIFY
         * @see NPOI.ss.usermodel.VerticalAlignment
         */
        public void SetVerticalAlignment(short align)
        {
            GetCellAlignment().Vertical = (VerticalAlignment)align;
        }


        /**
         * Gets border color
         *
         * @param side the border side
         * @return the used color
         */
        public XSSFColor GetBorderColor(BorderSide side)
        {
            switch (side)
            {
                case BorderSide.BOTTOM:
                    return BottomBorderXSSFColor;
                case BorderSide.RIGHT:
                    return RightBorderXSSFColor;
                case BorderSide.TOP:
                    return TopBorderXSSFColor;
                case BorderSide.LEFT:
                    return LeftBorderXSSFColor;
                default:
                    throw new ArgumentException("Unknown border: " + side);
            }
        }

        /**
         * Set the color to use for the selected border
         *
         * @param side - where to apply the color defInition
         * @param color - the color to use
         */
        public void SetBorderColor(BorderSide side, XSSFColor color)
        {
            switch (side)
            {
                case BorderSide.BOTTOM:
                    SetBottomBorderColor(color);
                    break;
                case BorderSide.RIGHT:
                    SetRightBorderColor(color);
                    break;
                case BorderSide.TOP:
                    SetTopBorderColor(color);
                    break;
                case BorderSide.LEFT:
                    SetLeftBorderColor(color);
                    break;
            }
        }
        private int FontId
        {
            get
            {
                if (_cellXf.IsSetFontId())
                {
                    return (int)_cellXf.fontId;
                }
                return (int)_cellStyleXf.fontId;
            }
        }

        /**
         * Get the cellAlignment object to use for manage alignment
         * @return XSSFCellAlignment - cell alignment
         */
        internal XSSFCellAlignment GetCellAlignment()
        {
            if (this._cellAlignment == null)
            {
                this._cellAlignment = new XSSFCellAlignment(GetCTCellAlignment());
            }
            return this._cellAlignment;
        }

        /**
         * Return the CT_CellAlignment instance for alignment
         *
         * @return CT_CellAlignment
         */
        internal CT_CellAlignment GetCTCellAlignment()
        {
            if (_cellXf.alignment == null)
            {
                _cellXf.alignment = new CT_CellAlignment();
            }
            return _cellXf.alignment;
        }

        /**
         * Returns a hash code value for the object. The hash is derived from the underlying CT_Xf bean.
         *
         * @return the hash code value for this style
         */
        public override int GetHashCode()
        {
            return _cellXf.ToString().GetHashCode();
        }

        /**
         * Checks is the supplied style is equal to this style
         *
         * @param o the style to check
         * @return true if the supplied style is equal to this style
         */
        public override bool Equals(Object o)
        {
            if (o == null || !(o is XSSFCellStyle)) return false;

            XSSFCellStyle cf = (XSSFCellStyle)o;
            return _cellXf.ToString().Equals(cf.GetCoreXf().ToString());
        }

        /**
         * Make a copy of this style. The underlying CT_Xf bean is Cloned,
         * the references to Fills and borders remain.
         *
         * @return a copy of this style
         */
        public Object Clone()
        {
            CT_Xf xf = (CT_Xf)_cellXf.Copy();

            int xfSize = _stylesSource.StyleXfsSize;
            int indexXf = _stylesSource.PutCellXf(xf);
            return new XSSFCellStyle(indexXf - 1, xfSize - 1, _stylesSource, _theme);
        }

        #region ICellStyle Members


        public IFont GetFont(IWorkbook parentWorkbook)
        {
            return this.GetFont();
        }

        public bool ShrinkToFit
        {
            get
            {
               CT_CellAlignment align = _cellXf.alignment;
               return align != null && align.shrinkToFit;
            }
            set
            {
               GetCTCellAlignment().shrinkToFit = value;
            }
        }

        public short BorderDiagonalColor
        {
            get
            {
                XSSFColor clr = DiagonalBorderXSSFColor;
                return clr == null ? IndexedColors.Black.Index : clr.Indexed;
            }
            set
            {
                XSSFColor clr = new XSSFColor();
                clr.Indexed = (value);
                SetDiagonalBorderColor(clr);
            }
        }

        public BorderStyle BorderDiagonalLineStyle
        {
            get
            {
                if (!_cellXf.applyBorder) return BorderStyle.None;

                int idx = (int)_cellXf.borderId;
                CT_Border ct = _stylesSource.GetBorderAt(idx).GetCTBorder();
                if (!ct.IsSetDiagonal())
                {
                    return BorderStyle.None;
                }
                else
                {
                    return (BorderStyle)ct.diagonal.style;
                }
            }
            set
            {
                CT_Border ct = GetCTBorder();
                CT_BorderPr pr = ct.IsSetDiagonal() ? ct.diagonal : ct.AddNewDiagonal();
                if (value == BorderStyle.None)
                    ct.unsetDiagonal();
                else
                    pr.style = (ST_BorderStyle)value;

                int idx = _stylesSource.PutBorder(new XSSFCellBorder(ct, _theme));

                _cellXf.borderId = (uint)idx;
                _cellXf.applyBorder = (true);
            }
        }

        public BorderDiagonal BorderDiagonal
        {
            get
            {
                CT_Border ct = GetCTBorder();
                if (ct.diagonalDown == true && ct.diagonalUp == true)
                    return BorderDiagonal.Both;
                else if (ct.diagonalDown == true)
                    return BorderDiagonal.Backward;
                else if (ct.diagonalUp == true)
                    return BorderDiagonal.Forward;
                else
                    return BorderDiagonal.None;
            }
            set
            {
                CT_Border ct = GetCTBorder();
                if (value == BorderDiagonal.Both)
                {
                    ct.diagonalDown = true;
                    ct.diagonalDownSpecified = true;
                    ct.diagonalUp = true;
                    ct.diagonalUpSpecified = true;
                }
                else if (value == BorderDiagonal.Forward)
                {
                    ct.diagonalDown = false;
                    ct.diagonalDownSpecified = false;
                    ct.diagonalUp = true;
                    ct.diagonalUpSpecified = true;
                }
                else if (value == BorderDiagonal.Backward)
                {
                    ct.diagonalDown = true;
                    ct.diagonalDownSpecified = true;
                    ct.diagonalUp = false;
                    ct.diagonalUpSpecified = false;
                }
                else
                {
                    ct.unsetDiagonal();
                    ct.diagonalDown = false;
                    ct.diagonalDownSpecified = false;
                    ct.diagonalUp = false;
                    ct.diagonalUpSpecified = false;
                }
            }
        }

        #endregion
    }

}
