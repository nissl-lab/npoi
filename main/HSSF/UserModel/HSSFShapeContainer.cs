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
    using System.Collections.Generic;

    /// <summary>
    /// An interface that indicates whether a class can contain children.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public interface HSSFShapeContainer : IEnumerable<HSSFShape>
    {
        /// <summary>
        /// Gets Any children contained by this shape.
        /// </summary>
        /// <value>The children.</value>
        IList<HSSFShape> Children { get; }

        /// <summary>
        /// dd shape to the list of child records
        /// </summary>
        /// <param name="shape">shape</param>
        void AddShape(HSSFShape shape);

        /// <summary>
        /// set coordinates of this group relative to the parent
        /// </summary>
        /// <param name="x1">x1</param>
        /// <param name="y1">y1</param>
        /// <param name="x2">x2</param>
        /// <param name="y2">y2</param>
        void SetCoordinates(int x1, int y1, int x2, int y2);

        void Clear();

        /// <summary>
        /// Get the top left x coordinate of this group.
        /// </summary>
        int X1 { get; }

        /// <summary>
        /// Get the top left y coordinate of this group.
        /// </summary>
        int Y1 { get; }

        /// <summary>
        /// Get the bottom right x coordinate of this group.
        /// </summary>
        int X2 { get; }

        /// <summary>
        /// Get the bottom right y coordinate of this group.
        /// </summary>
        int Y2 { get; }

        /**
         * remove first level shapes
         * @param shape to be removed
         * @return true if shape is removed else return false
         */
        bool RemoveShape(HSSFShape shape);
    }
}
