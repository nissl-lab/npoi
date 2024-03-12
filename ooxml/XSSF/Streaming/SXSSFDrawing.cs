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

namespace NPOI.XSSF.Streaming
{
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;

    /// <summary>
    /// Streaming version of Drawing.
    /// Delegates most tasks to the non-streaming XSSF code.
    /// TODO: Potentially, Comment and Chart need a similar streaming wrapper like Picture.
    /// </summary>
    public class SXSSFDrawing : IDrawing
    {
        private SXSSFWorkbook _wb;
        private XSSFDrawing _drawing;

        public SXSSFDrawing(SXSSFWorkbook workbook, XSSFDrawing Drawing)
        {
            this._wb = workbook;
            this._drawing = Drawing;
        }
        public IPicture CreatePicture(IClientAnchor anchor, int pictureIndex)
        {
            XSSFPicture pict = (XSSFPicture)_drawing.CreatePicture(anchor, pictureIndex);
            return new SXSSFPicture(_wb, pict);
        }
        public IComment CreateCellComment(IClientAnchor anchor)
        {
            return _drawing.CreateCellComment(anchor);
        }
        public IChart CreateChart(IClientAnchor anchor)
        {
            return _drawing.CreateChart(anchor);
        }
        public IClientAnchor CreateAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2)
        {
            return _drawing.CreateAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
        }
        //public ObjectData CreateObjectData(IClientAnchor anchor, int storageId, int pictureIndex)
        //{
        //    return _drawing.CreateObjectData(anchor, storageId, pictureIndex);
        //}
        //public Iterator<XSSFShape> iterator()
        //{
        //    return _drawing.Shapes.Iterator();
        //}

    }

}

