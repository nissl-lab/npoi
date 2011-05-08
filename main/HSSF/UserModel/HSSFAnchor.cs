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


    /// <summary>
    /// An anchor Is what specifics the position of a shape within a client object
    /// or within another containing shape.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public abstract class HSSFAnchor
    {
        int dx1;
        int dy1;
        int dx2;
        int dy2;

        public HSSFAnchor()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFAnchor"/> class.
        /// </summary>
        /// <param name="dx1">The DX1.</param>
        /// <param name="dy1">The dy1.</param>
        /// <param name="dx2">The DX2.</param>
        /// <param name="dy2">The dy2.</param>
        public HSSFAnchor(int dx1, int dy1, int dx2, int dy2)
        {
            this.dx1 = dx1;
            this.dy1 = dy1;
            this.dx2 = dx2;
            this.dy2 = dy2;
        }

        /// <summary>
        /// Gets or sets the DX1.
        /// </summary>
        /// <value>The DX1.</value>
        public int Dx1
        { 
            get{return dx1;}
            set { this.dx1 = value; }
        }
        /// <summary>
        /// Gets or sets the dy1.
        /// </summary>
        /// <value>The dy1.</value>
        public int Dy1
        {
            get { return dy1; }
            set { this.dy1 = value; }
        }
        /// <summary>
        /// Gets or sets the dy2.
        /// </summary>
        /// <value>The dy2.</value>
        public int Dy2
        {
            get { return dy2; }
            set { this.dy2 = value; }
        }
        /// <summary>
        /// Gets or sets the DX2.
        /// </summary>
        /// <value>The DX2.</value>
        public int Dx2
        {
            get { return dx2; }
            set { this.dx2 = value; }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is horizontally flipped.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is horizontally flipped; otherwise, <c>false</c>.
        /// </value>
        public abstract bool IsHorizontallyFlipped { get; }
        /// <summary>
        /// Gets a value indicating whether this instance is vertically flipped.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is vertically flipped; otherwise, <c>false</c>.
        /// </value>
        public abstract bool IsVerticallyFlipped { get; }
    }
}