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
        /// Move and Resize With Anchor Cells
        /// </para>
        /// <para>
        /// Specifies that the current drawing shall move and
        /// resize to maintain its row and column anchors (i.e. the
        /// object is anchored to the actual from and to row and column)
        /// </para>
        /// </summary>
        MoveAndResize = 0,

        /// <summary>
        /// <para>
        /// Move With Cells but Do Not Resize
        /// </para>
        /// <para>
        /// Specifies that the current drawing shall move with its
        /// row and column (i.e. the object is anchored to the
        /// actual from row and column), but that the size shall remain absolute.
        /// </para>
        /// <para>
        /// 
        /// </para>
        /// <para>
        /// If Additional rows/columns are Added between the from and to locations of the drawing,
        /// the drawing shall move its to anchors as needed to maintain this same absolute size.
        /// </para>
        /// </summary>
        MoveDontResize = 2,

        /// <summary>
        /// <para>
        /// Do Not Move or Resize With Underlying Rows/Columns
        /// </para>
        /// <para>
        /// Specifies that the current start and end positions shall
        /// be maintained with respect to the distances from the
        /// absolute start point of the worksheet.
        /// </para>
        /// <para>
        /// 
        /// </para>
        /// <para>
        /// If Additional rows/columns are Added before the
        /// drawing, the drawing shall move its anchors as needed
        /// to maintain this same absolute position.
        /// </para>
        /// </summary>
        DontMoveAndResize = 3

    }

    /// <summary>
    /// A client anchor is attached to an excel worksheet.  It anchors against a
    /// top-left and bottom-right cell.
    /// </summary>
    /// @author Yegor Kozlov
    public interface IClientAnchor
    {

        /// <summary>
        /// Returns the column (0 based) of the first cell.
        /// </summary>
        /// <returns>0-based column of the first cell.</returns>
        int Col1 { get; set; }

        /// <summary>
        /// Returns the column (0 based) of the second cell.
        /// </summary>
        /// <returns>0-based column of the second cell.</returns>
        int Col2 { get; set; }

        /// <summary>
        /// Returns the row (0 based) of the first cell.
        /// </summary>
        /// <returns>0-based row of the first cell.</returns>
        int Row1 { get; set; }


        /// <summary>
        /// Returns the row (0 based) of the second cell.
        /// </summary>
        /// <returns>0-based row of the second cell.</returns>
        int Row2 { get; set; }


        /// <summary>
        /// Returns the x coordinate within the first cell
        /// </summary>
        /// <returns>the x coordinate within the first cell</returns>
        int Dx1 { get; set; }


        /// <summary>
        /// Returns the y coordinate within the first cell
        /// </summary>
        /// <returns>the y coordinate within the first cell</returns>
        int Dy1 { get; set; }


        /// <summary>
        /// Sets the y coordinate within the second cell
        /// </summary>
        /// <returns>the y coordinate within the second cell</returns>
        int Dy2 { get; set; }

        /// <summary>
        /// Returns the x coordinate within the second cell
        /// </summary>
        /// <returns>the x coordinate within the second cell</returns>
        int Dx2 { get; set; }


        /// <summary>
        /// <para>
        /// s the anchor type
        /// </para>
        /// <para>
        /// 0 = Move and size with Cells, 2 = Move but don't size with cells, 3 = Don't move or size with cells.
        /// </para>
        /// </summary>
        /// <returns>the anchor type</returns>
        /// <see cref="MOVE_AND_RESIZE" />
        /// <see cref="MOVE_DONT_RESIZE" />
        /// <see cref="DONT_MOVE_AND_RESIZE" />
        AnchorType AnchorType { get; set; }

    }
}