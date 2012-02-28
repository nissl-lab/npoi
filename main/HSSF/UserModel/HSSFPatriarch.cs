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
    using System.Collections;
    using NPOI.DDF;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    using NPOI.SS.UserModel;
    using System.Collections.Generic;

    /// <summary>
    /// The patriarch is the toplevel container for shapes in a sheet.  It does
    /// little other than act as a container for other shapes and Groups.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class HSSFPatriarch : HSSFShapeContainer,IDrawing
    {
        List<HSSFShape> _shapes = new List<HSSFShape>();
        public HSSFSheet _sheet;
        int x1 = 0;
        int y1 = 0;
        int x2 = 1023;
        int y2 = 255;

        /**
         * The EscherAggregate we have been bound to.
         * (This will handle writing us out into records,
         *  and building up our shapes from the records)
         */
        private EscherAggregate _boundAggregate;

        /// <summary>
        /// Creates the patriarch.
        /// </summary>
        /// <param name="sheet">the sheet this patriarch is stored in.</param>
        /// <param name="boundAggregate">The bound aggregate.</param>
        public HSSFPatriarch(HSSFSheet sheet, EscherAggregate boundAggregate)
        {
            this._boundAggregate = boundAggregate;
            this._sheet = sheet;
        }

        /// <summary>
        /// Creates a new Group record stored Under this patriarch.
        /// </summary>
        /// <param name="anchor">the client anchor describes how this Group is attached
        /// to the sheet.</param>
        /// <returns>the newly created Group.</returns>
        public HSSFShapeGroup CreateGroup(HSSFClientAnchor anchor)
        {
            HSSFShapeGroup group = new HSSFShapeGroup(null, anchor);
            group.Anchor = anchor;
            AddShape(group);
            return group;
        }

        /// <summary>
        /// Creates a simple shape.  This includes such shapes as lines, rectangles,
        /// and ovals.
        /// </summary>
        /// <param name="anchor">the client anchor describes how this Group is attached
        /// to the sheet.</param>
        /// <returns>the newly created shape.</returns>
        public HSSFSimpleShape CreateSimpleShape(HSSFClientAnchor anchor)
        {
            HSSFSimpleShape shape = new HSSFSimpleShape(null, anchor);
            shape.Anchor = anchor;
            AddShape(shape);
            return shape;
        }

        /// <summary>
        /// Creates a picture.
        /// </summary>
        /// <param name="anchor">the client anchor describes how this Group is attached
        /// to the sheet.</param>
        /// <param name="pictureIndex">Index of the picture.</param>
        /// <returns>the newly created shape.</returns>
        public IPicture CreatePicture(HSSFClientAnchor anchor, int pictureIndex)
        {
            HSSFPicture shape = new HSSFPicture(null, (HSSFClientAnchor)anchor);
            shape.PictureIndex=pictureIndex;
            shape.Anchor = (HSSFClientAnchor)anchor;
            AddShape(shape);

            EscherBSERecord bse = (_sheet.Workbook as HSSFWorkbook).Workbook.GetBSERecord(pictureIndex);
            bse.Ref = (bse.Ref + 1);
            return shape;
        }
        public IPicture CreatePicture(IClientAnchor anchor, int pictureIndex)
        {
            return CreatePicture((HSSFClientAnchor)anchor, pictureIndex);
        }

        /// <summary>
        /// Creates a polygon
        /// </summary>
        /// <param name="anchor">the client anchor describes how this Group is attached
        /// to the sheet.</param>
        /// <returns>the newly Created shape.</returns>
        public HSSFPolygon CreatePolygon(IClientAnchor anchor)
        {
            HSSFPolygon shape = new HSSFPolygon(null, (HSSFAnchor)anchor);
            shape.Anchor = (HSSFAnchor)anchor;
            AddShape(shape);
            return shape;
        }

        /// <summary>
        /// Constructs a textbox Under the patriarch.
        /// </summary>
        /// <param name="anchor">the client anchor describes how this Group is attached
        /// to the sheet.</param>
        /// <returns>the newly Created textbox.</returns>
        public HSSFSimpleShape CreateTextbox(IClientAnchor anchor)
        {
            HSSFTextbox shape = new HSSFTextbox(null, (HSSFAnchor)anchor);
            shape.Anchor = (HSSFAnchor)anchor;
            AddShape(shape);
            return shape;
        }
        /**
         * Constructs a cell comment.
         *
         * @param anchor    the client anchor describes how this comment is attached
         *                  to the sheet.
         * @return      the newly created comment.
         */
        public HSSFComment CreateComment(HSSFAnchor anchor)
        {
            HSSFComment shape = new HSSFComment(null, anchor);
            shape.Anchor = anchor;
            AddShape(shape);
            return shape;
        }
        /**
         * YK: used to create autofilters
         *
         * @see org.apache.poi.hssf.usermodel.HSSFSheet#setAutoFilter(int, int, int, int)
         */
        public HSSFSimpleShape CreateComboBox(HSSFAnchor anchor)
        {
            HSSFSimpleShape shape = new HSSFSimpleShape(null, anchor);
            shape.ShapeType = HSSFSimpleShape.OBJECT_TYPE_COMBO_BOX;
            shape.Anchor = anchor;
            AddShape(shape);
            return shape;
        }

        /// <summary>
        /// Constructs a cell comment.
        /// </summary>
        /// <param name="anchor">the client anchor describes how this comment is attached
        /// to the sheet.</param>
        /// <returns>the newly created comment.</returns>
        public IComment CreateCellComment(IClientAnchor anchor)
        {
            return CreateComment((HSSFAnchor)anchor);
        }

        /// <summary>
        /// Returns a list of all shapes contained by the patriarch.
        /// </summary>
        /// <value>The children.</value>
        public System.Collections.IList Children
        {
            get { return _shapes; }
        }

        /**
         * add a shape to this drawing
         */
        internal void AddShape(HSSFShape shape)
        {
            shape._patriarch = this;
            _shapes.Add(shape);
        }
        /// <summary>
        /// Total count of all children and their children's children.
        /// </summary>
        /// <value>The count of all children.</value>
        public int CountOfAllChildren
        {
            get
            {
                int count = _shapes.Count;
                for (IEnumerator iterator = _shapes.GetEnumerator(); iterator.MoveNext(); )
                {
                    HSSFShape shape = (HSSFShape)iterator.Current;
                    count += shape.CountOfAllChildren;
                }
                return count;
            }
        }
        /// <summary>
        /// Sets the coordinate space of this Group.  All children are contrained
        /// to these coordinates.
        /// </summary>
        /// <param name="x1">The x1.</param>
        /// <param name="y1">The y1.</param>
        /// <param name="x2">The x2.</param>
        /// <param name="y2">The y2.</param>
        public void SetCoordinates(int x1, int y1, int x2, int y2)
        {
            this.x1 = x1;
            this.y1 = y1;
            this.x2 = x2;
            this.y2 = y2;
        }

        /// <summary>
        /// Does this HSSFPatriarch contain a chart?
        /// (Technically a reference to a chart, since they
        /// Get stored in a different block of records)
        /// FIXME - detect chart in all cases (only seems
        /// to work on some charts so far)
        /// </summary>
        /// <returns>
        /// 	<c>true</c> if this instance contains chart; otherwise, <c>false</c>.
        /// </returns>
        public bool ContainsChart()
        {
            // TODO - support charts properly in usermodel

            // We're looking for a EscherOptRecord
            EscherOptRecord optRecord = (EscherOptRecord)
                _boundAggregate.FindFirstWithId(EscherOptRecord.RECORD_ID);
            if (optRecord == null)
            {
                // No opt record, can't have chart
                return false;
            }

            for (IEnumerator it = optRecord.EscherProperties.GetEnumerator(); it.MoveNext(); )
            {
                EscherProperty prop = (EscherProperty)it.Current;
                if (prop.PropertyNumber == 896 && prop.IsComplex)
                {
                    EscherComplexProperty cp = (EscherComplexProperty)prop;
                    String str = StringUtil.GetFromUnicodeLE(cp.ComplexData);
                    //Console.Error.WriteLine(str);
                    if (str.Equals("Chart 1\0"))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// The top left x coordinate of this Group.
        /// </summary>
        /// <value>The x1.</value>
        public int X1
        {
            get { return x1; }
        }

        /// <summary>
        /// The top left y coordinate of this Group.
        /// </summary>
        /// <value>The y1.</value>
        public int Y1
        {
            get { return y1; }
        }

        /// <summary>
        /// The bottom right x coordinate of this Group.
        /// </summary>
        /// <value>The x2.</value>
        public int X2
        {
            get { return x2; }
        }

        /// <summary>
        /// The bottom right y coordinate of this Group.
        /// </summary>
        /// <value>The y2.</value>
        public int Y2
        {
            get { return y2; }
        }

        /// <summary>
        /// Returns the aggregate escher record we're bound to
        /// </summary>
        /// <returns></returns>
        protected EscherAggregate _GetBoundAggregate()
        {
            return _boundAggregate;
        }

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
        public IClientAnchor CreateAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2)
        {
            return new HSSFClientAnchor(dx1, dy1, dx2, dy2, (short)col1, row1, (short)col2, row2);
        }

        public IChart CreateChart(IClientAnchor anchor)
        {
            throw new RuntimeException("NotImplemented");
        }
    }
}