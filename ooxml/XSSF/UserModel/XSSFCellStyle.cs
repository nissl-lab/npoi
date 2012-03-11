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
         * @param cellXfId The main XF for the cell
         * @param cellStyleXfId Optional, style xf
         * @param stylesSource Styles Source to work off
         */
        public XSSFCellStyle(int cellXfId, int cellStyleXfId, StylesTable stylesSource, ThemesTable theme)
        {
            _cellXfId = cellXfId;
            _stylesSource = stylesSource;
            _cellXf = stylesSource.GetCellXfAt(this._cellXfId);
            _cellStyleXf = stylesSource.GetCellStyleXfAt(cellStyleXfId);
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


        /**
         * Creates an empty Cell Style
         */
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
                    _cellXf = src.GetCoreXf();
                    _cellStyleXf = src.GetStyleXf();
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
                            _cellXf.unsetExtLst();

                        // Create a new Xf with the same contents
                        _cellXf = CT_Xf.Parse(
                              src.GetCoreXf().ToString()
                        );
                        // Swap it over
                        _stylesSource.ReplaceCellXfAt(_cellXfId, _cellXf);
                    }
                    catch (XmlException e)
                    {
                        throw new POIXMLException(e);
                    }

                    // Copy the format
                    String fmt = src.GetDataFormatString();
                    DataFormat=(
                          (new XSSFDataFormat(_stylesSource)).GetFormat(fmt)
                    );

                    // Copy the font
                    try
                    {
                        CT_Font ctFont = new CT_Font(
                              src.GetFont().GetCTFont().ToString()
                        );
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

        public HorizontalAlignment Alignment
        {
            get
            {
                return GetAlignmentEnum();
            }
            set 
            {
                GetCellAlignment().SetHorizontal(value);
            }
        }

        /**
         * Get the type of horizontal alignment for the cell
         *
         * @return HorizontalAlignment - the type of alignment
         * @see NPOI.ss.usermodel.HorizontalAlignment
         */
        public HorizontalAlignment GetAlignmentEnum()
        {
            CT_CellAlignment align = _cellXf.alignment;
            if (align != null && align.IsSetHorizontal())
            {
                return (HorizontalAlignment)(align.horizontal - 1);
            }
            return HorizontalAlignment.GENERAL;
        }

        public BorderStyle BorderBottom
        {
            get
            {
                if (!_cellXf.applyBorder) return BorderStyle.NONE;

                int idx = (int)_cellXf.borderId;
                CT_Border ct = _stylesSource.GetBorderAt(idx).GetCTBorder();
                ST_BorderStyle? ptrn = ct.IsSetBottom() ? ct.bottom.style : null;
                return ptrn == null ? BorderStyle.NONE : (BorderStyle)(ptrn - 1);
            }
            set 
            {
                CT_Border ct = GetCTBorder();
                CT_BorderPr pr = ct.IsSetBottom() ? ct.bottom : ct.AddNewBottom();
                if (value == BorderStyle.NONE) ct.unsetBottom();
                else pr.style=(ST_BorderStyle)(value + 1);

                int idx = _stylesSource.PutBorder(new XSSFCellBorder(ct, _theme));

                _cellXf.borderId = (uint)idx;
                _cellXf.applyBorder = (true);
            }
        }


        public BorderStyle BorderLeft
        {
            get
            {
                if (!_cellXf.applyBorder) return BorderStyle.NONE;

                int idx = (int)_cellXf.borderId;
                CT_Border ct = _stylesSource.GetBorderAt(idx).GetCTBorder();
                ST_BorderStyle? ptrn = ct.IsSetLeft() ? ct.left.style : null;
                return ptrn == null ? BorderStyle.NONE : (BorderStyle)(ptrn - 1);
            }
            set 
            {
                CT_Border ct = GetCTBorder();
                CT_BorderPr pr = ct.IsSetLeft() ? ct.left : ct.AddNewLeft();
                if (value == BorderStyle.NONE) ct.unsetLeft();
                else pr.style=(ST_BorderStyle)(value + 1);

                int idx = _stylesSource.PutBorder(new XSSFCellBorder(ct, _theme));

                _cellXf.borderId = (uint)idx;
                _cellXf.applyBorder = (true);

            }
        }


        /**
         * Get the type of border to use for the right border of the cell
         *
         * @return short - border type, default value is {@link NPOI.ss.usermodel.CellStyle#BorderStyle.NONE}
         * @see NPOI.ss.usermodel.CellStyle#BorderStyle.NONE
         * @see NPOI.ss.usermodel.CellStyle#BORDER_THIN
         * @see NPOI.ss.usermodel.CellStyle#BORDER_MEDIUM
         * @see NPOI.ss.usermodel.CellStyle#BORDER_DASHED
         * @see NPOI.ss.usermodel.CellStyle#BORDER_DOTTED
         * @see NPOI.ss.usermodel.CellStyle#BORDER_THICK
         * @see NPOI.ss.usermodel.CellStyle#BORDER_DOUBLE
         * @see NPOI.ss.usermodel.CellStyle#BORDER_HAIR
         * @see NPOI.ss.usermodel.CellStyle#BORDER_MEDIUM_DASHED
         * @see NPOI.ss.usermodel.CellStyle#BORDER_DASH_DOT
         * @see NPOI.ss.usermodel.CellStyle#BORDER_MEDIUM_DASH_DOT
         * @see NPOI.ss.usermodel.CellStyle#BORDER_DASH_DOT_DOT
         * @see NPOI.ss.usermodel.CellStyle#BORDER_MEDIUM_DASH_DOT_DOT
         * @see NPOI.ss.usermodel.CellStyle#BORDER_SLANTED_DASH_DOT
         */
        public BorderStyle BorderRight
        {
            get
            {
                if (!_cellXf.applyBorder) return BorderStyle.NONE;

                int idx = (int)_cellXf.borderId;
                CT_Border ct = _stylesSource.GetBorderAt(idx).GetCTBorder();
                ST_BorderStyle? ptrn = ct.IsSetRight() ? ct.right.style : null;
                return ptrn == null ? BorderStyle.NONE : (BorderStyle)(ptrn - 1);
            }
            set 
            {
                CT_Border ct = GetCTBorder();
                CT_BorderPr pr = ct.IsSetRight() ? ct.right : ct.AddNewRight();
                if (value == BorderStyle.NONE) ct.unsetRight();
                else pr.style = (ST_BorderStyle)((int)value + 1);

                int idx = _stylesSource.PutBorder(new XSSFCellBorder(ct, _theme));

                _cellXf.borderId = (uint)idx;
                _cellXf.applyBorder = (true);
            }
        }

        public BorderStyle BorderTop
        {
            get
            {
                if (!_cellXf.applyBorder) return BorderStyle.NONE;

                int idx = (int)_cellXf.borderId;
                CT_Border ct = _stylesSource.GetBorderAt(idx).GetCTBorder();
                ST_BorderStyle? ptrn = ct.IsSetTop() ? ct.top.style : null;
                return ptrn == null ? BorderStyle.NONE : (BorderStyle)(ptrn - 1);
            }
            set 
            {
                CT_Border ct = GetCTBorder();
                CT_BorderPr pr = ct.IsSetTop() ? ct.top : ct.AddNewTop();
                if (value == BorderStyle.NONE) ct.unsetTop();
                else pr.style = (ST_BorderStyle)(value + 1);

                int idx = _stylesSource.PutBorder(new XSSFCellBorder(ct, _theme));

                _cellXf.borderId = (uint)idx;
                _cellXf.applyBorder = (true);
            }
        }

        /**
         * Get the color to use for the bottom border
         * <br/>
         * Color is optional. When missing, IndexedColors.AUTOMATIC is implied.
         * @return the index of the color defInition, default value is {@link NPOI.ss.usermodel.IndexedColors#AUTOMATIC}
         * @see NPOI.ss.usermodel.IndexedColors
         */
        public short BottomBorderColor
        {
            get
            {
                XSSFColor clr = GetBottomBorderXSSFColor();
                return clr == null ? IndexedColors.BLACK.Index : clr.GetIndexed();
            }
            set 
            {
                XSSFColor clr = new XSSFColor();
                clr.SetIndexed(value);
                SetBottomBorderColor(clr);
            }
        }

        /**
         * Get the color to use for the bottom border as a {@link XSSFColor}
         *
         * @return the used color or <code>null</code> if not Set
         */
        public XSSFColor GetBottomBorderXSSFColor()
        {
            if (!_cellXf.applyBorder) return null;

            int idx = (int)_cellXf.borderId;
            XSSFCellBorder border = _stylesSource.GetBorderAt(idx);

            return border.GetBorderColor(BorderSide.BOTTOM);
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
                _cellXf.applyNumberFormat = (true);
                _cellXf.numFmtId = (uint)value;
            }
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

        /**
         * Get the background fill color.
         * <p>
         * Note - many cells are actually Filled with a foreground
         *  Fill, not a background fill - see {@link #getFillForegroundColor()}
         * </p>
         * @return fill color, default value is {@link NPOI.ss.usermodel.IndexedColors#AUTOMATIC}
         * @see NPOI.ss.usermodel.IndexedColors
         */
        public short FillBackgroundColor
        {
            get
            {
                XSSFColor clr = FillBackgroundXSSFColor;
                return clr == null ? IndexedColors.AUTOMATIC.Index : clr.GetIndexed();
            }
            set 
            {
                XSSFColor clr = new XSSFColor();
                clr.SetIndexed(value);
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
        public XSSFColor FillBackgroundXSSFColor
        {
            get
            {
                if (!_cellXf.applyFill) return null;

                int FillIndex = (int)_cellXf.fillId;
                XSSFCellFill fg = _stylesSource.GetFillAt(FillIndex);

                XSSFColor FillBackgroundColor = fg.GetFillBackgroundColor();
                if (FillBackgroundColor != null && _theme != null)
                {
                    _theme.InheritFromThemeAsRequired(FillBackgroundColor);
                }
                return FillBackgroundColor;
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
                XSSFColor clr = this.FillForegroundXSSFColor;
                return clr == null ? IndexedColors.AUTOMATIC.Index : clr.GetIndexed();
            }
            set 
            {
                XSSFColor clr = new XSSFColor();
                clr.SetIndexed(value);
                SetFillForegroundColor(clr);
            }
        }
        /**
         * Get the foreground fill color.
         *
         * @return XSSFColor - fill color or <code>null</code> if not Set
         */
        public XSSFColor FillForegroundXSSFColor
        {
            get
            {
                if (!_cellXf.applyFill) return null;

                int FillIndex = (int)_cellXf.fillId;
                XSSFCellFill fg = _stylesSource.GetFillAt(FillIndex);

                XSSFColor FillForegroundColor = fg.GetFillForegroundColor();
                if (FillForegroundColor != null && _theme != null)
                {
                    _theme.InheritFromThemeAsRequired(FillForegroundColor);
                }
                return FillForegroundColor;
            }
        }


        public FillPatternType FillPattern
        {
            get
            {
                if (!_cellXf.applyFill) return 0;

                int FillIndex = (int)_cellXf.fillId;
                XSSFCellFill fill = _stylesSource.GetFillAt(FillIndex);

                ST_PatternType ptrn = fill.GetPatternType();
                if (ptrn == null) return FillPatternType.NO_FILL;
                return (FillPatternType)(ptrn - 1);
            }
            set 
            {
                CT_Fill ct = GetCTFill();
                CT_PatternFill ptrn = ct.IsSetPatternFill() ? ct.GetPatternFill() : ct.AddNewPatternFill();
                if (value == FillPatternType.NO_FILL && ptrn.IsSetPatternType()) 
                    ptrn.unsetPatternType();
                else ptrn.patternType = (ST_PatternType)(value + 1);

                int idx = _stylesSource.PutFill(new XSSFCellFill(ct));

                _cellXf.fillId = (uint)idx;
                _cellXf.applyFill = (true);
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
                _font = _stylesSource.GetFontAt(GetFontId());
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
                return (short)GetFontId();
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
                GetCellAlignment().SetIndent(value);
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
                XSSFColor clr = GetLeftBorderXSSFColor();
                return clr == null ? IndexedColors.BLACK.Index : clr.GetIndexed();
            }
            set 
            {
                XSSFColor clr = new XSSFColor();
                clr.SetIndexed(value);
                SetLeftBorderColor(clr);
            }
        }

        /**
         * Get the color to use for the left border
         *
         * @return the index of the color defInition or <code>null</code> if not Set
         * @see NPOI.ss.usermodel.IndexedColors
         */
        public XSSFColor GetLeftBorderXSSFColor()
        {
            if (!_cellXf.applyBorder) return null;

            int idx = (int)_cellXf.borderId;
            XSSFCellBorder border = _stylesSource.GetBorderAt(idx);

            return border.GetBorderColor(BorderSide.LEFT);
        }

        /**
         * Get whether the cell's using this style are locked
         *
         * @return whether the cell using this style are locked
         */
        public bool IsLocked
        {
            get
            {

                if (!_cellXf.IsSetProtection() || !_cellXf.protection.IsSetLocked())
                {
                    return true;
                }
                return (bool)_cellXf.protection.locked;
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

        /**
         * Get the color to use for the right border
         *
         * @return the index of the color defInition, default value is {@link NPOI.ss.usermodel.IndexedColors#BLACK}
         * @see NPOI.ss.usermodel.IndexedColors
         */
        public short RightBorderColor
        {
            get
            {
                XSSFColor clr = GetRightBorderXSSFColor();
                return clr == null ? IndexedColors.BLACK.Index : clr.GetIndexed();
            }
            set 
            {
                XSSFColor clr = new XSSFColor();
                clr.SetIndexed(value);
                SetRightBorderColor(clr);
            }
        }
        /**
         * Get the color to use for the right border
         *
         * @return the used color or <code>null</code> if not Set
         */
        public XSSFColor GetRightBorderXSSFColor()
        {
            if (!_cellXf.applyBorder) return null;

            int idx = (int)_cellXf.borderId;
            XSSFCellBorder border = _stylesSource.GetBorderAt(idx);

            return border.GetBorderColor(BorderSide.RIGHT);
        }

        /**
         * Get the degree of rotation for the text in the cell
         * <p>
         * Expressed in degrees. Values range from 0 to 180. The first letter of
         * the text is considered the center-point of the arc.
         * <br/>
         * For 0 - 90, the value represents degrees above horizon. For 91-180 the degrees below the
         * horizon is calculated as:
         * <br/>
         * <code>[degrees below horizon] = 90 - textRotation.</code>
         * </p>
         *
         * @return rotation degrees (between 0 and 180 degrees)
         */
        public short Rotation
        {
            get
            {
                CT_CellAlignment align = _cellXf.alignment;
                return (short)(align == null ? 0 : align.textRotation);
            }
            set 
            {
                GetCellAlignment().SetTextRotation(value);
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
                XSSFColor clr = GetTopBorderXSSFColor();
                return clr == null ? IndexedColors.BLACK.Index : clr.GetIndexed();
            }
            set 
            {
                XSSFColor clr = new XSSFColor();
                clr.SetIndexed(value);
                SetTopBorderColor(clr);
            }
        }

        /**
         * Get the color to use for the top border
         *
         * @return the used color or <code>null</code> if not Set
         */
        public XSSFColor GetTopBorderXSSFColor()
        {
            if (!_cellXf.applyBorder) return null;

            int idx = (int)_cellXf.borderId;
            XSSFCellBorder border = _stylesSource.GetBorderAt(idx);

            return border.GetBorderColor(BorderSide.TOP);
        }

        /**
         * Get the type of vertical alignment for the cell
         *
         * @return align the type of alignment, default value is {@link NPOI.ss.usermodel.CellStyle#VERTICAL_BOTTOM}
         * @see NPOI.ss.usermodel.CellStyle#VERTICAL_TOP
         * @see NPOI.ss.usermodel.CellStyle#VERTICAL_CENTER
         * @see NPOI.ss.usermodel.CellStyle#VERTICAL_BOTTOM
         * @see NPOI.ss.usermodel.CellStyle#VERTICAL_JUSTIFY
         */
        public VerticalAlignment VerticalAlignment
        {
            get
            {
                return GetVerticalAlignmentEnum();
            }
            set 
            {
                GetCellAlignment().SetVertical(value);
            }
        }

        /**
         * Get the type of vertical alignment for the cell
         *
         * @return the type of alignment, default value is {@link NPOI.ss.usermodel.VerticalAlignment#BOTTOM}
         * @see NPOI.ss.usermodel.VerticalAlignment
         */
        public VerticalAlignment GetVerticalAlignmentEnum()
        {
            CT_CellAlignment align = _cellXf.alignment;
            if (align != null && align.IsSetVertical())
            {
                return (VerticalAlignment)(align.vertical - 1);
            }
            return VerticalAlignment.BOTTOM;
        }

        /**
         * Whether the text should be wrapped
         *
         * @return  a bool value indicating if the text in a cell should be line-wrapped within the cell.
         */
        public bool WrapText
        {
            get
            {
                CT_CellAlignment align = _cellXf.alignment;
                return align != null && align.wrapText;
            }
            set 
            {
                GetCellAlignment().SetWrapText(value);
            }
        }



        /**
         * Set the type of border to use for the bottom border of the cell
         *
         * @param border - type of border to use
         * @see NPOI.ss.usermodel.BorderStyle
         */
        public void SetBorderBottom(BorderStyle border)
        {
            SetBorderBottom(border);
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
         * <i>Note: Ensure Foreground color is Set prior to background</i>
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
         * It is necessary to Set the fill style in order
         * for the color to be Shown in the cell.
         *
         * @param color - the color to use
         */
        public void SetFillBackgroundColor(XSSFColor color)
        {
            CT_Fill ct = GetCTFill();
            CT_PatternFill ptrn = ct.GetPatternFill();
            if (color == null)
            {
                if (ptrn != null) ptrn.unsetBgColor();
            }
            else
            {

                if (ptrn == null) ptrn = ct.AddNewPatternFill();
                ptrn.bgColor = (color.GetCTColor());
            }

            int idx = _stylesSource.PutFill(new XSSFCellFill(ct));

            _cellXf.fillId = (uint)idx;
            _cellXf.applyFill= (true);
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
                if (ptrn != null) ptrn.unsetFgColor();
            }
            else
            {
                if (ptrn == null) ptrn = ct.AddNewPatternFill();
                ptrn.fgColor =(color.GetCTColor());
            }

            int idx = _stylesSource.PutFill(new XSSFCellFill(ct));

            _cellXf.fillId = (uint)idx;
            _cellXf.applyFill = (true);
        }

        /**
         * Get a <b>copy</b> of the currently used CT_Fill, if none is used, return a new instance.
         */
        private CT_Fill GetCTFill()
        {
            CT_Fill ct;
            if (_cellXf.applyFill)
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
        private CT_Border GetCTBorder()
        {
            CT_Border ct;
            if (_cellXf.applyBorder)
            {
                int idx = (int)_cellXf.borderId;
                XSSFCellBorder cf = _stylesSource.GetBorderAt(idx);

                ct = (CT_Border)cf.GetCTBorder().Copy();
            }
            else
            {
                ct = new CT_Border();
            }
            return ct;
        }

        /**
         * This element is used to specify cell fill information for pattern and solid color cell Fills. For solid cell Fills (no pattern),
         * foreground color is used is used. For cell Fills with patterns specified, then the cell fill color is specified by the background color element.
         *
         * @param ptrn the fill pattern to use
         * @see #setFillBackgroundColor(short)
         * @see #setFillForegroundColor(short)
         * @see NPOI.ss.usermodel.FillPatternType
         */
        public void SetFillPattern(FillPatternType ptrn)
        {
            SetFillPattern(ptrn);
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
                this._cellXf.applyFont = (true);
            }
            else
            {
                this._cellXf.applyFont = (false);
            }
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
            GetCellAlignment().SetVertical((VerticalAlignment)align);
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
                    return GetBottomBorderXSSFColor();
                case BorderSide.RIGHT:
                    return GetRightBorderXSSFColor();
                case BorderSide.TOP:
                    return GetTopBorderXSSFColor();
                case BorderSide.LEFT:
                    return GetLeftBorderXSSFColor();
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
        private int GetFontId()
        {
            if (_cellXf.IsSetFontId())
            {
                return (int)_cellXf.fontId;
            }
            return (int)_cellStyleXf.fontId;
        }

        /**
         * Get the cellAlignment object to use for manage alignment
         * @return XSSFCellAlignment - cell alignment
         */
        protected XSSFCellAlignment GetCellAlignment()
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
        private CT_CellAlignment GetCTCellAlignment()
        {
            if (_cellXf.alignment == null)
            {
                _cellXf.alignment=new CT_CellAlignment();
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

            int xfSize = _stylesSource._GetStyleXfsSize();
            int indexXf = _stylesSource.PutCellXf(xf);
            return new XSSFCellStyle(indexXf - 1, xfSize - 1, _stylesSource, _theme);
        }

        #region ICellStyle Members


        public IFont GetFont(IWorkbook parentWorkbook)
        {
            throw new NotImplementedException();
        }

        public bool ShrinkToFit
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        #endregion
    }

}