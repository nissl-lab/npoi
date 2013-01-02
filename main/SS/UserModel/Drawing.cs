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

    /**
     * @author Yegor Kozlov
     */
    public interface IDrawing
    {
        /**
      * Creates a picture.
      * @param anchor       the client anchor describes how this picture is
      *                     attached to the sheet.
      * @param pictureIndex the index of the picture in the workbook collection
      *                     of pictures.
      *
      * @return the newly created picture.
      */
        IPicture CreatePicture(IClientAnchor anchor, int pictureIndex);

        /**
         * Creates a comment.
         * @param anchor the client anchor describes how this comment is attached
         *               to the sheet.
         * @return the newly created comment.
         */
        IComment CreateCellComment(IClientAnchor anchor);

        /**
         * Creates a chart.
         * @param anchor the client anchor describes how this chart is attached to
         *               the sheet.
         * @return the newly created chart
         */
        IChart CreateChart(IClientAnchor anchor);

        /**
         * Creates a new client anchor and sets the top-left and bottom-right
         * coordinates of the anchor.
         *
         * @param dx1  the x coordinate in EMU within the first cell.
         * @param dy1  the y coordinate in EMU within the first cell.
         * @param dx2  the x coordinate in EMU within the second cell.
         * @param dy2  the y coordinate in EMU within the second cell.
         * @param col1 the column (0 based) of the first cell.
         * @param row1 the row (0 based) of the first cell.
         * @param col2 the column (0 based) of the second cell.
         * @param row2 the row (0 based) of the second cell.
         * @return the newly created client anchor
         */
        IClientAnchor CreateAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2);
    }

}