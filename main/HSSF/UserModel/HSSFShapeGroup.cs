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

    /// <summary>
    /// A shape Group may contain other shapes.  It was no actual form on the
    /// sheet.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class HSSFShapeGroup : HSSFShape, HSSFShapeContainer
    {
        IList shapes = new ArrayList();
        int x1 = 0;
        int y1 = 0;
        int x2 = 1023;
        int y2 = 255;


        public HSSFShapeGroup(HSSFShape parent, HSSFAnchor anchor):base(parent, anchor)
        {
            
        }

        /// <summary>
        /// Create another Group Under this Group.
        /// </summary>
        /// <param name="anchor">the position of the new Group.</param>
        /// <returns>the Group</returns>
        public HSSFShapeGroup CreateGroup(HSSFChildAnchor anchor)
        {
            HSSFShapeGroup group = new HSSFShapeGroup(this, anchor);
            group.Anchor = anchor;
            shapes.Add(group);
            return group;
        }

        /// <summary>
        /// Create a new simple shape Under this Group.
        /// </summary>
        /// <param name="anchor">the position of the shape.</param>
        /// <returns>the shape</returns>
        public HSSFSimpleShape CreateShape(HSSFChildAnchor anchor)
        {
            HSSFSimpleShape shape = new HSSFSimpleShape(this, anchor);
            shape.Anchor = anchor;
            shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Create a new textbox Under this Group.
        /// </summary>
        /// <param name="anchor">the position of the shape.</param>
        /// <returns>the textbox</returns>
        public HSSFTextbox CreateTextbox(HSSFChildAnchor anchor)
        {
            HSSFTextbox shape = new HSSFTextbox(this, anchor);
            shape.Anchor = anchor;
            shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Creates a polygon
        /// </summary>
        /// <param name="anchor">the client anchor describes how this Group Is attached
        /// to the sheet.</param>
        /// <returns>the newly Created shape.</returns>
        public HSSFPolygon CreatePolygon(HSSFChildAnchor anchor)
        {
            HSSFPolygon shape = new HSSFPolygon(this, anchor);
            shape.Anchor = anchor;
            shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Creates a picture.
        /// </summary>
        /// <param name="anchor">the client anchor describes how this Group Is attached
        /// to the sheet.</param>
        /// <param name="pictureIndex">Index of the picture.</param>
        /// <returns>the newly Created shape.</returns>
        public HSSFPicture CreatePicture(HSSFChildAnchor anchor, int pictureIndex)
        {
            HSSFPicture shape = new HSSFPicture(this, anchor);
            shape.Anchor = anchor;
            shape.PictureIndex=pictureIndex;
            shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Return all children contained by this shape.
        /// </summary>
        /// <value></value>
        public System.Collections.IList Children
        {
            get { return shapes; }
        }

        /// <summary>
        /// Sets the coordinate space of this Group.  All children are constrained
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
        /// Gets The top left x coordinate of this Group.
        /// </summary>
        /// <value>The x1.</value>
        public int X1
        {
            get { return x1; }
        }

        /// <summary>
        /// Gets The top left y coordinate of this Group.
        /// </summary>
        /// <value>The y1.</value>
        public int Y1
        {
            get
            {
                return y1;
            }
        }

        /// <summary>
        /// Gets The bottom right x coordinate of this Group.
        /// </summary>
        /// <value>The x2.</value>
        public int X2
        {
            get
            {
                return x2;
            }
        }

        /// <summary>
        /// Gets the bottom right y coordinate of this Group.
        /// </summary>
        /// <value>The y2.</value>
        public int Y2
        {
            get
            {
                return y2;
            }
        }

        /// <summary>
        /// Count of all children and their childrens children.
        /// </summary>
        /// <value></value>
        public override int CountOfAllChildren
        {
            get
            {
                int count = shapes.Count;
                for (IEnumerator iterator = shapes.GetEnumerator(); iterator.MoveNext(); )
                {
                    HSSFShape shape = (HSSFShape)iterator.Current;
                    count += shape.CountOfAllChildren;
                }
                return count;
            }
        }
    }
}