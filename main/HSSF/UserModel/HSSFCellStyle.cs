/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


namespace NPOI.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Util;
    using NPOI.SS.UserModel;
    using System.Collections.Generic;

    /// <summary>
    /// High level representation of the style of a cell in a sheet of a workbook.
    /// @author  Andrew C. Oliver (acoliver at apache dot org)
    /// @author Jason Height (jheight at chariot dot net dot au)
    /// </summary>
    public class HSSFCellStyle : ICellStyle
    {
        private ExtendedFormatRecord _format = null;
        private short index = 0;
        private NPOI.HSSF.Model.InternalWorkbook _workbook = null;


        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFCellStyle"/> class.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <param name="rec">The record.</param>
        /// <param name="workbook">The workbook.</param>
        public HSSFCellStyle(short index, ExtendedFormatRecord rec, HSSFWorkbook workbook)
            :this(index, rec, workbook.Workbook)
        {
            
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFCellStyle"/> class.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <param name="rec">The record.</param>
        /// <param name="workbook">The workbook.</param>
        public HSSFCellStyle(short index, ExtendedFormatRecord rec, NPOI.HSSF.Model.InternalWorkbook workbook)
        {
            this._workbook = workbook;
            this.index = index;
            _format = rec;
        }

        /// <summary>
        /// Get the index within the HSSFWorkbook (sequence within the collection of ExtnededFormat objects)
        /// </summary>
        /// <value>Unique index number of the Underlying record this style represents (probably you don't care
        /// Unless you're comparing which one is which)</value>
        public short Index
        {
            get { return index; }
        }
        /// <summary>
        /// Gets the parent style.
        /// </summary>
        /// <value>the parent style for this cell style.
        /// In most cases this will be null, but in a few
        /// cases there'll be a fully defined parent.</value>
        public HSSFCellStyle ParentStyle
        {
            get
            {
                short parentIndex = _format.ParentIndex;
                // parentIndex equal 0xFFF indicates no inheritance from a cell style XF (See 2.4.353 XF)
                if ( parentIndex == 0|| parentIndex == 0xFFF)
                {
                    return null;
                }
                return new HSSFCellStyle(
                        parentIndex,
                        _workbook.GetExFormatAt(parentIndex),
                        _workbook
                );
            }
        }
        /// <summary>
        /// Get the index of the format
        /// </summary>
        /// <value>The data format.</value>
        public short DataFormat
        {
            get { return _format.FormatIndex; }
            set { _format.FormatIndex = (value); }
        }
        private static short lastDateFormat = short.MinValue;
        private static List<FormatRecord> lastFormats = null;
        private static String getDataFormatStringCache = null;
        /// <summary>
        /// Get the contents of the format string, by looking up
        /// the DataFormat against the bound workbook
        /// </summary>
        /// <returns></returns>
        public String GetDataFormatString()
        {
            //HSSFDataFormat format = new HSSFDataFormat(workbook);
            //return format.GetFormat(DataFormat);

            if (getDataFormatStringCache != null)
            {
                if (lastDateFormat == DataFormat && _workbook.Formats.Equals(lastFormats))
                {
                    return getDataFormatStringCache;
                }
            }

            lastFormats = _workbook.Formats;
            lastDateFormat = DataFormat;

            getDataFormatStringCache = GetDataFormatString(_workbook);

            return getDataFormatStringCache;
        }

        /// <summary>
        /// Get the contents of the format string, by looking up the DataFormat against the supplied workbook
        /// </summary>
        /// <param name="workbook">The workbook</param>
        /// <returns>the format string or "General" if not found</returns>
        public String GetDataFormatString(IWorkbook workbook)
        {
            HSSFDataFormat format = new HSSFDataFormat(((HSSFWorkbook)workbook).Workbook);

            int idx = DataFormat;
            return idx == -1 ? "General" : format.GetFormat(DataFormat);
        }
        /// <summary>
        /// Get the contents of the format string, by looking up
        /// the DataFormat against the supplied workbook
        /// </summary>
        /// <param name="workbook">The internal workbook.</param>
        /// <returns></returns>
        public String GetDataFormatString(NPOI.HSSF.Model.InternalWorkbook workbook)
        {
            HSSFDataFormat format = new HSSFDataFormat(workbook);

            return format.GetFormat(DataFormat);
        }


        /// <summary>
        /// Set the font for this style
        /// </summary>
        /// <param name="font">a font object Created or retreived from the HSSFWorkbook object</param>
        public void SetFont(NPOI.SS.UserModel.IFont font)
        {
            _format.IsIndentNotParentFont=(true);
            short fontindex = font.Index;
            _format.FontIndex=(fontindex);
        }

        /// <summary>
        /// Gets the index of the font for this style.
        /// </summary>
        /// <value>The index of the font.</value>
        public short FontIndex
        {
            get { return _format.FontIndex; }
        }

        /// <summary>
        /// Gets the font for this style
        /// </summary>
        /// <param name="parentWorkbook">The parent workbook that this style belongs to.</param>
        /// <returns></returns>
        public IFont GetFont(IWorkbook parentWorkbook)
        {
            return ((HSSFWorkbook)parentWorkbook).GetFontAt(FontIndex);
        }

        /// <summary>
        /// Get whether the cell's using this style are to be hidden
        /// </summary>
        /// <value>whether the cell using this style should be hidden</value>
        public bool IsHidden
        {
            get { return _format.IsHidden; }
            set
            {
                _format.IsIndentNotParentCellOptions=(true);
                _format.IsHidden=(value);
            }
        }


        /// <summary>
        /// Get whether the cell's using this style are to be locked
        /// </summary>
        /// <value>whether the cell using this style should be locked</value>
        public bool IsLocked
        {
            get { return _format.IsLocked; }
            set
            {
                _format.IsIndentNotParentCellOptions=(true);
                _format.IsLocked=(value);
            }
        }

        /// <summary>
        /// Get the type of horizontal alignment for the cell
        /// </summary>
        /// <value> the type of alignment</value>
        public HorizontalAlignment Alignment
        {
            get { return (HorizontalAlignment)_format.Alignment; }
            set
            {
                _format.IsIndentNotParentAlignment=(true);
                _format.Alignment=(short)value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the text should be wrapped
        /// </summary>
        /// <value><c>true</c> if [wrap text]; otherwise, <c>false</c>.</value>
        public bool WrapText
        {
            get
            {
                return _format.WrapText;
            }
            set
            {
                _format.IsIndentNotParentAlignment = (true);
                _format.WrapText = (value);
            }
        }

        /// <summary>
        /// Gets or sets the vertical alignment for the cell.
        /// </summary>
        /// <value>the type of alignment</value>
        public VerticalAlignment VerticalAlignment
        {
            get
            {
                return (VerticalAlignment)_format.VerticalAlignment;

            }
            set { _format.VerticalAlignment=(short)value; }
        }

        /// <summary>
        /// Gets or sets the degree of rotation for the text in the cell
        /// </summary>
        /// <value>The rotation degrees (between -90 and 90 degrees).</value>
        public short Rotation
        {
            get
            {
                short rotation = _format.Rotation;
                if (rotation == 0xff) 
                {
                    return rotation;
                }
                if (rotation > 90)
                    //This is actually the 4th quadrant
                    rotation = (short)(90 - rotation);
                return rotation;
            }
            set
            {
                short rotation = value;

                if (rotation == 0xff) 
                {

                }else if ((value < 0) && (value >= -90))
                {
                    //Take care of the funny 4th quadrant Issue
                    //The 4th quadrant (-1 to -90) is stored as (91 to 180)
                    rotation = (short)(90 - value);
                }

                else if ((value < -90) || (value > 90))
                {
                    //Do not allow an incorrect rotation to be Set
                    throw new ArgumentException("The rotation must be between -90 and 90 degrees, or 0xff");
                }
                _format.Rotation=(rotation);
            }
        }
        /// <summary>
        /// Verifies that this style belongs to the supplied Workbook.
        /// Will throw an exception if it belongs to a different one.
        /// This is normally called when trying to assign a style to a
        /// cell, to ensure the cell and the style are from the same
        /// workbook (if they're not, it won't work)
        /// </summary>
        /// <param name="wb">The workbook.</param>
        public void VerifyBelongsToWorkbook(HSSFWorkbook wb)
        {
            if (wb.Workbook != _workbook)
            {
                throw new ArgumentException("This Style does not belong to the supplied Workbook. Are you trying to assign a style from one workbook to the cell of a differnt workbook?");
            }
        }


        /// <summary>
        /// Gets or sets the number of spaces to indent the text in the cell
        /// </summary>
        /// <value>number of spaces</value>
        public short Indention
        {
            get{return _format.Indent;}
            set { _format.Indent = (value); }
        }

        /// <summary>
        /// Gets or sets the type of border to use for the left border of the cell
        /// </summary>
        /// <value>The border type.</value>
        public BorderStyle BorderLeft
        {
            get { return (BorderStyle)_format.BorderLeft; }
            set
            {
                _format.IsIndentNotParentBorder=(true);
                _format.BorderLeft=(short)value;
            }
        }

        /// <summary>
        /// Gets or sets the type of border to use for the right border of the cell
        /// </summary>
        /// <value>The border type.</value>
        public BorderStyle BorderRight
        {
            get { return (BorderStyle)_format.BorderRight; }
            set
            {
                _format.IsIndentNotParentBorder = (true);
                _format.BorderRight = (short)value;
            }
        }

        /// <summary>
        /// Gets or sets the type of border to use for the top border of the cell
        /// </summary>
        /// <value>The border type.</value>
        public BorderStyle BorderTop
        {
            get { return (BorderStyle)_format.BorderTop; }
            set
            {
                _format.IsIndentNotParentBorder = (true);
                _format.BorderTop = (short)value;
            }
        }

        /// <summary>
        /// Gets or sets the type of border to use for the bottom border of the cell
        /// </summary>
        /// <value>The border type.</value>
        public BorderStyle BorderBottom
        {
            get { return (BorderStyle)_format.BorderBottom; }
            set
            {
                _format.IsIndentNotParentBorder = true;
                _format.BorderBottom = (short)value;
            }
        }

        /// <summary>
        /// Gets or sets the color to use for the left border
        /// </summary>
        /// <value>The index of the color definition</value>
        public short LeftBorderColor
        {
            get { return _format.LeftBorderPaletteIdx; }
            set { _format.LeftBorderPaletteIdx = (value); }
        }


        /// <summary>
        /// Gets or sets the color to use for the left border.
        /// </summary>
        /// <value>The index of the color definition</value>
        public short RightBorderColor
        {
            get { return _format.RightBorderPaletteIdx; }
            set { _format.RightBorderPaletteIdx = (value); }
        }


        /// <summary>
        /// Gets or sets the color to use for the top border
        /// </summary>
        /// <value>The index of the color definition.</value>
        public short TopBorderColor
        {
            get { return _format.TopBorderPaletteIdx; }
            set { _format.TopBorderPaletteIdx = (value); }
        }


        /// <summary>
        /// Gets or sets the color to use for the left border
        /// </summary>
        /// <value>The index of the color definition.</value>
        public short BottomBorderColor
        {
            get { return _format.BottomBorderPaletteIdx; }
            set { _format.BottomBorderPaletteIdx = (value); }
        }

        /// <summary>
        /// Gets or sets the color to use for the diagional border
        /// </summary>
        /// <value>The index of the color definition.</value>
        public short BorderDiagonalColor
        {
            get { return _format.AdtlDiagBorderPaletteIdx; }
            set { _format.AdtlDiagBorderPaletteIdx = value; }
        }

        /// <summary>
        /// Gets or sets the line type  to use for the diagional border
        /// </summary>
        /// <value>The line type.</value>
        public BorderStyle BorderDiagonalLineStyle
        {
            get { return (BorderStyle)_format.AdtlDiagLineStyle; }
            set { _format.AdtlDiagLineStyle=(short)value; }
        }

        /// <summary>
        /// Gets or sets the type of diagional border
        /// </summary>.
        /// <value>The border diagional type.</value>
        public BorderDiagonal BorderDiagonal
        {
            get { return (BorderDiagonal)_format.Diagonal; }
            set { _format.Diagonal = (short)value; }
        }

        /// <summary>
        /// Gets or sets whether the cell is shrink-to-fit
        /// </summary>
        public bool ShrinkToFit
        {
            get { return _format.ShrinkToFit; }
            set { _format.ShrinkToFit = value; }
        }

        /**
         * Get or set the reading order, for RTL/LTR ordering of
         *  the text.
         * <p>0 means Context (Default), 1 means Left To Right,
         *  and 2 means Right to Left</p>
         *
         * @return order - the reading order (0,1,2)
         */
        public short ReadingOrder
        {
            get
            {
                return _format.ReadingOrder;
            }
            set
            {
                _format.ReadingOrder = (value);
            }
        }

        /// <summary>
        /// Gets or sets the fill pattern. - Set to 1 to Fill with foreground color
        /// </summary>
        /// <value>The fill pattern.</value>
        public FillPattern FillPattern
        {
            get
            {
                return (FillPattern)_format.AdtlFillPattern;
            }
            set { _format.AdtlFillPattern=(short)value; }
        }

        /// <summary>
        /// Checks if the background and foreground Fills are Set correctly when one
        /// or the other is Set to the default color.
        /// Works like the logic table below:
        /// BACKGROUND   FOREGROUND
        /// NONE         AUTOMATIC
        /// 0x41         0x40
        /// NONE         RED/ANYTHING
        /// 0x40         0xSOMETHING
        /// </summary>
        private void CheckDefaultBackgroundFills()
        {
            if (_format.FillForeground == HSSFColor.Automatic.Index)
            {
                //JMH: Why +1, hell why not. I guess it made some sense to someone at the time. Doesnt
                //to me now.... But experience has shown that when the fore is Set to AUTOMATIC then the
                //background needs to be incremented......
                if (_format.FillBackground != (HSSFColor.Automatic.Index + 1))
                    FillBackgroundColor = HSSFColor.Automatic.Index + 1;
            }
            else if (_format.FillBackground == HSSFColor.Automatic.Index + 1)
                //Now if the forground Changes to a non-AUTOMATIC color the background Resets itself!!!
                if (_format.FillForeground != HSSFColor.Automatic.Index)
                    FillBackgroundColor=(HSSFColor.Automatic.Index);
        }
        /**
         * Clones all the style information from another
         *  HSSFCellStyle, onto this one. This
         *  HSSFCellStyle will then have all the same
         *  properties as the source, but the two may
         *  be edited independently.
         * Any stylings on this HSSFCellStyle will be lost!
         *
         * The source HSSFCellStyle could be from another
         *  HSSFWorkbook if you like. This allows you to
         *  copy styles from one HSSFWorkbook to another.
         */
        public void CloneStyleFrom(ICellStyle source)
        {
            if (source is HSSFCellStyle)
            {
                this.CloneStyleFrom((HSSFCellStyle)source);
            }
            else
            {
                throw new ArgumentException("Can only clone from one HSSFCellStyle to another, not between HSSFCellStyle and XSSFCellStyle");
            }

        }
        /// <summary>
        /// Clones all the style information from another
        /// HSSFCellStyle, onto this one. This
        /// HSSFCellStyle will then have all the same
        /// properties as the source, but the two may
        /// be edited independently.
        /// Any stylings on this HSSFCellStyle will be lost!
        /// The source HSSFCellStyle could be from another
        /// HSSFWorkbook if you like. This allows you to
        /// copy styles from one HSSFWorkbook to another.
        /// </summary>
        /// <param name="source">The source.</param>
        public void CloneStyleFrom(HSSFCellStyle source)
        {
            // First we need to clone the extended format
            //  record
            _format.CloneStyleFrom(source._format);

            // Handle matching things if we cross workbooks
            if (_workbook != source._workbook)
            {
                lastDateFormat = short.MinValue;
                lastFormats = null;
                getDataFormatStringCache = null;
                // Then we need to clone the format string,
                //  and update the format record for this
                short fmt = (short)_workbook.CreateFormat(
                        source.GetDataFormatString()
                );
                this.DataFormat=(fmt);

                // Finally we need to clone the font,
                //  and update the format record for this
                FontRecord fr = _workbook.CreateNewFont();
                fr.CloneStyleFrom(
                        source._workbook.GetFontRecordAt(
                                source.FontIndex
                        )
                );

                HSSFFont font = new HSSFFont(
                        (short)_workbook.GetFontIndex(fr), fr
                );
                this.SetFont(font);
            }
        }
        /// <summary>
        /// Gets or sets the color of the fill background.
        /// </summary>
        /// <value>The color of the fill background.</value>
        /// Set the background Fill color.
        /// <example>
        /// cs.SetFillPattern(HSSFCellStyle.FINE_DOTS );
        /// cs.SetFillBackgroundColor(new HSSFColor.RED().Index);
        /// optionally a Foreground and background Fill can be applied:
        /// Note: Ensure Foreground color is Set prior to background
        /// cs.SetFillPattern(HSSFCellStyle.FINE_DOTS );
        /// cs.SetFillForegroundColor(new HSSFColor.BLUE().Index);
        /// cs.SetFillBackgroundColor(new HSSFColor.RED().Index);
        /// or, for the special case of SOLID_Fill:
        /// cs.SetFillPattern(HSSFCellStyle.SOLID_FOREGROUND );
        /// cs.SetFillForegroundColor(new HSSFColor.RED().Index);
        /// It is necessary to Set the Fill style in order
        /// for the color to be shown in the cell.
        /// </example>
        public short FillBackgroundColor
        {
            get
            {
                short result = _format.FillBackground;
                //JMH: Do this ridiculous conversion, and let HSSFCellStyle
                //internally migrate back and forth
                if (result == (HSSFColor.Automatic.Index + 1))
                    return HSSFColor.Automatic.Index;
                else return result;
            }
            set
            {
                _format.FillBackground=value;
                CheckDefaultBackgroundFills();
            }
        }
        public IColor FillBackgroundColorColor
        {
            get
            {
                HSSFPalette pallette = new HSSFPalette(_workbook.CustomPalette);
                return pallette.GetColor(FillBackgroundColor);
            }
        }
        /// <summary>
        /// Gets or sets the foreground Fill color
        /// </summary>
        /// <value>Fill color.</value>
        /// @see org.apache.poi.hssf.usermodel.HSSFPalette#GetColor(short)
        public short FillForegroundColor
        {
            get{return _format.FillForeground;}
            set
            {
                _format.FillForeground = value;
                CheckDefaultBackgroundFills();
            }
        }
        public IColor FillForegroundColorColor
        {
            get
            {
                HSSFPalette pallette = new HSSFPalette(_workbook.CustomPalette);
                return pallette.GetColor(FillForegroundColor);
            }
        }
        /**
 * Gets the name of the user defined style.
 * Returns null for built in styles, and
 *  styles where no name has been defined
 */
        public String UserStyleName
        {
            get
            {
                StyleRecord sr = _workbook.GetStyleRecord(index);
                if (sr == null)
                {
                    return null;
                }
                if (sr.IsBuiltin)
                {
                    return null;
                }
                return sr.Name;
            }
            set 
            {
                StyleRecord sr = _workbook.GetStyleRecord(index);
                if (sr == null)
                {
                    sr = _workbook.CreateStyleRecord(index);
                }
                // All Style records start as "builtin", but generally
                //  only 20 and below really need to be
                if (sr.IsBuiltin && index <= 20)
                {
                    throw new ArgumentException("Unable to set user specified style names for built in styles!");
                }
                sr.Name  = value;
            }
        }


        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="T:System.Object"/>.
        /// </returns>
        public override int GetHashCode()
        {
            int prime = 31;
            int result = 1;
            result = prime * result + ((_format == null) ? 0 : _format.GetHashCode());
            result = prime * result + index;
            return result;
        }

        /// <summary>
        /// Determines whether the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <param name="obj">The <see cref="T:System.Object"/> to compare with the current <see cref="T:System.Object"/>.</param>
        /// <returns>
        /// true if the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>; otherwise, false.
        /// </returns>
        /// <exception cref="T:System.NullReferenceException">
        /// The <paramref name="obj"/> parameter is null.
        /// </exception>
        public override bool Equals(Object obj)
        {
            if (this == obj) return true;
            if (obj == null) return false;
            if (obj is HSSFCellStyle)
            {
                HSSFCellStyle other = (HSSFCellStyle)obj;
                if (_format == null)
                {
                    if (other._format != null)
                        return false;
                }
                else if (!_format.Equals(other._format))
                    return false;
                if (index != other.index)
                    return false;
                return true;
            }
            return false;
        }
    }
}