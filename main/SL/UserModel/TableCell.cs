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

using System;
using System.Drawing;

namespace NPOI.SL.UserModel
{
	public enum BorderEdge { bottom, left, top, right }
	public interface TableCell<S, P>: TextShape<S, P>
		where S : Shape<S, P>
		where P : TextParagraph<S, P, TextRun>
	{
		

		/**
		 * Return line style of given edge or {@code null} if border is not defined
		 *
		 * @param edge the border edge
		 * @return line style of given edge or {@code null} if border is not defined
		 */
		StrokeStyle getBorderStyle(BorderEdge edge);

		/**
		 * Sets the {@link StrokeStyle} of the given border edge.
		 * A {@code null} property of the style is ignored.
		 *
		 * @param edge border edge
		 * @param style the new stroke style
		 */
		void setBorderStyle(BorderEdge edge, StrokeStyle style);

		/**
		 * Convenience method for setting the border width.
		 *
		 * @param edge border edge
		 * @param width the new border width
		 */
		void setBorderWidth(BorderEdge edge, double width);

		/**
		 * Convenience method for setting the border color.
		 *
		 * @param edge border edge
		 * @param color the new border color
		 */
		void setBorderColor(BorderEdge edge, Color color);

		/**
		 * Convenience method for setting the border line compound.
		 *
		 * @param edge border edge
		 * @param compound the new border line compound
		 */
		//void setBorderCompound(BorderEdge edge, LineCompound compound);

		/**
		 * Convenience method for setting the border line dash.
		 *
		 * @param edge border edge
		 * @param dash the new border line dash
		 */
		//void setBorderDash(BorderEdge edge, LineDash dash);

		/**
		 * Remove all line attributes of the given border edge
		 *
		 * @param edge the border edge to be cleared
		 */
		void removeBorder(BorderEdge edge);

		/**
		 * Get the number of columns to be spanned/merged
		 *
		 * @return the grid span
		 * 
		 * @since POI 3.15-beta2
		 */
		int getGridSpan();

		/**
		 * Get the number of rows to be spanned/merged
		 *
		 * @return the row span
		 * 
		 * @since POI 3.15-beta2
		 */
		int getRowSpan();

		/**
		 * Return if this cell is part of a merged cell. The top/left cell of a merged region is not regarded as merged -
		 * its grid and/or row span is greater than one. 
		 *
		 * @return true if this a merged cell
		 * 
		 * @since POI 3.15-beta2
		 */
		bool isMerged();
	}
}