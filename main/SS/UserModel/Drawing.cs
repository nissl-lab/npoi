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
    /// <summary>
    /// High level representation of spreadsheet Drawing.
    /// </summary>
    public interface IDrawing<out T> : IShapeContainer<T> where T : class, IShape
    {
        /// <summary>
        /// Creates a picture.
        /// </summary>
        /// <param name="anchor">      the client anchor describes how this picture is
        /// attached to the sheet.
        /// </param>
        /// <param name="pictureIndex">the index of the picture in the workbook collection
        /// of pictures.
        /// </param>
        /// <return>the newly created picture.</return>
        IPicture CreatePicture(IClientAnchor anchor, int pictureIndex);

        /// <summary>
        /// Creates a comment.
        /// </summary>
        /// <param name="anchor">the client anchor describes how this comment is attached
        /// to the sheet.
        /// </param>
        /// <return>the newly created comment.</return>
        IComment CreateCellComment(IClientAnchor anchor);

        /// <summary>
        /// Creates a chart.
        /// </summary>
        /// <param name="anchor">the client anchor describes how this chart is attached to
        /// the sheet.
        /// </param>
        /// <return>the newly created chart</return>
        IChart CreateChart(IClientAnchor anchor);

        /// <summary>
        /// Creates a new client anchor and Sets the top-left and bottom-right
        /// coordinates of the anchor.
        /// </summary>
        /// <param name="dx1"> the x coordinate in EMU within the first cell.</param>
        /// <param name="dy1"> the y coordinate in EMU within the first cell.</param>
        /// <param name="dx2"> the x coordinate in EMU within the second cell.</param>
        /// <param name="dy2"> the y coordinate in EMU within the second cell.</param>
        /// <param name="col1">the column (0 based) of the first cell.</param>
        /// <param name="row1">the row (0 based) of the first cell.</param>
        /// <param name="col2">the column (0 based) of the second cell.</param>
        /// <param name="row2">the row (0 based) of the second cell.</param>
        /// <return>the newly created client anchor</return>
        IClientAnchor CreateAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2);
    }
}