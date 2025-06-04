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
namespace NPOI.SS.UserModel
{
    public enum AnchorType : int
    {
        /// <summary>
        /// <para>
        /// Move and Resize With Anchor Cells (0)
        /// </para>
        /// <para>
        /// Specifies that the current Drawing shall Move and
        /// resize to maintain its row and column anchors (i.e. the
        /// object is anchored to the actual from and to row and column)
        /// </para>
        /// </summary>
        MoveAndResize = 0,

        /// <summary>
        /// <para>
        /// Don't Move but do Resize With Anchor Cells (1)
        /// </para>
        /// <para>
        /// Specifies that the current Drawing shall not Move with its
        /// row and column, but should be resized. This option is not normally
        /// used, but is included for completeness.
        /// </para>
        /// <para>
        /// Note: Excel has no Setting for this combination, nor does the ECMA standard.
        /// </para>
        /// </summary>
        DontMoveDoResize =1,

        /// <summary>
        /// <para>
        /// Move With Cells but Do Not Resize (2)
        /// </para>
        /// <para>
        /// Specifies that the current Drawing shall Move with its
        /// row and column (i.e. the object is anchored to the
        /// actual from row and column), but that the size shall remain absolute.
        /// </para>
        /// <para>
        /// If additional rows/columns are added between the from and to locations of the Drawing,
        /// the Drawing shall Move its to anchors as needed to maintain this same absolute size.
        /// </para>
        /// </summary>
        MoveDontResize = 2,

        /// <summary>
        /// <para>
        /// Do Not Move or Resize With Underlying Rows/Columns (3)
        /// </para>
        /// <para>
        /// Specifies that the current start and end positions shall
        /// be maintained with respect to the distances from the
        /// absolute start point of the worksheet.
        /// </para>
        /// <para>
        /// If additional rows/columns are added before the
        /// Drawing, the Drawing shall Move its anchors as needed
        /// to maintain this same absolute position.
        /// </para>
        /// </summary>
        DontMoveAndResize = 3
    }

    /// <summary>
    /// A client anchor is attached to an excel worksheet.  It anchors against
    /// absolute coordinates, a top-left cell and fixed height and width, or
    /// a top-left and bottom-right cell, depending on the <see cref="AnchorType"/>:
    /// <list type="number">
    /// <item><description> <see cref="AnchorType.DontMoveAndResize" /> == absolute top-left coordinates and width/height, no cell references</description></item>
    /// <item><description> <see cref="AnchorType.MoveDontResize" /> == fixed top-left cell reference, absolute width/height</description></item>
    /// <item><description> <see cref="AnchorType.MoveAndResize" /> == fixed top-left and bottom-right cell references, dynamic width/height</description></item>
    /// </list>
    /// Note this class only reports the current values for possibly calculated positions and sizes.
    /// If the sheet row/column sizes or positions shift, this needs updating via external calculations.
    /// </summary>
    public interface IClientAnchor
    {
                /// <summary>
        /// Get or set the column (0 based) of the first cell, or -1 if there is no top-left anchor cell.
        /// This is the case for absolute positioning <see cref="AnchorType.MoveAndResize" />
        /// </summary>
        /// <return>0-based column of the first cell or -1 if none.</return>
        int Col1 { get; set; }

        /// <summary>
        /// Get or set the column (0 based) of the second cell, or -1 if there is no bottom-right anchor cell.
        /// This is the case for absolute positioning (<see cref="AnchorType.DontMoveAndResize" />)
        /// and absolute sizing (<see cref="AnchorType.MoveDontResize" />.
        /// </summary>
        /// <return>0-based column of the second cell or -1 if none.</return>
        int Col2 { get; set; }

        /// <summary>
        /// Get or set the row (0 based) of the first cell, or -1 if there is no bottom-right anchor cell.
        /// This is the case for absolute positioning (<see cref="AnchorType.DontMoveAndResize" />).
        /// </summary>
        /// <return>0-based row of the first cell or -1 if none.</return>
        int Row1 { get; set; }


        /// <summary>
        /// Get or set the row (0 based) of the second cell, or -1 if there is no bottom-right anchor cell.
        /// This is the case for absolute positioning (<see cref="AnchorType.DontMoveAndResize" />)
        /// and absolute sizing (<see cref="AnchorType.MoveDontResize" />.
        /// </summary>
        /// <return>0-based row of the second cell or -1 if none.</return>
        int Row2 { get; set; }


        /// <summary>
        /// <para>
        /// Get or set the x coordinate within the first cell.
        /// </para>
        /// <para>
        /// Note - XSSF and HSSF have a slightly different coordinate
        ///  system, values in XSSF are larger by a factor of
        ///  <see cref="NPOI.Util.Units.EMU_PER_PIXEL" />
        /// </para>
        /// </summary>
        /// <return>the x coordinate within the first cell</return>
        int Dx1 { get; set; }


        /// <summary>
        /// <para>
        /// Get or set the y coordinate within the first cell
        /// </para>
        /// <para>
        /// Note - XSSF and HSSF have a slightly different coordinate
        ///  system, values in XSSF are larger by a factor of
        ///  <see cref="NPOI.Util.Units.EMU_PER_PIXEL" />
        /// </para>
        /// </summary>
        /// <return>the y coordinate within the first cell</return>
        int Dy1 { get; set; }


        /// <summary>
        /// <para>
        /// Get or set the y coordinate within the second cell
        /// </para>
        /// <para>
        /// Note - XSSF and HSSF have a slightly different coordinate
        ///  system, values in XSSF are larger by a factor of
        ///  <see cref="NPOI.Util.Units.EMU_PER_PIXEL" />
        /// </para>
        /// </summary>
        /// <return>the y coordinate within the second cell</return>
        int Dy2 { get; set; }

        /// <summary>
        /// <para>
        /// Get or set the x coordinate within the second cell
        /// </para>
        /// <para>
        /// Note - XSSF and HSSF have a slightly different coordinate
        ///  system, values in XSSF are larger by a factor of
        ///  <see cref="NPOI.Util.Units.EMU_PER_PIXEL" />
        /// </para>
        /// </summary>
        /// <return>the x coordinate within the second cell</return>
        int Dx2 { get; set; }


        /// <summary>
        /// Get or set the anchor type
        /// Changed from returning an int to an enum in POI 3.14 beta 1.
        /// </summary>
        /// <return>the anchor type</return>
        AnchorType AnchorType { get; set; }

    }
}