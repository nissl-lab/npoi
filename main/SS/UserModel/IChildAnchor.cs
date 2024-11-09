/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
    /// Common interface for anchors.
    /// 
    /// An anchor is what specifics the position of a shape within a client object
    /// or within another containing shape.
    /// </summary>
    public interface IChildAnchor
    {
        /// <summary>
        /// get or set x coordinate of the left up corner
        /// </summary>
        int Dx1 { get; set; }

        /// <summary>
        /// get or set y coordinate of the left up corner
        /// </summary>
        int Dy1 {  get; set; }

        /// <summary>
        /// get or set x coordinate of the right down corner
        /// </summary>
        int Dx2 { get; set; }
        /// <summary>
        /// get or set y coordinate of the right down corner
        /// </summary>
        int Dy2 { get; set; }
    }
}
