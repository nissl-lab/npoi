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
namespace NPOI.Common.UserModel.Fonts
{
	using System;
	using System.Collections.Generic;
	using System.Text;

	/**
 * A property of a font that describes its general appearance.
 * 
 * @since POI 3.17-beta2
 */
	public enum FontFamilyEnum
	{
		/**
	     * The default font is specified, which is implementation-dependent.
	     */
		FF_DONTCARE = 0x00,
		/**
	     * Fonts with variable stroke widths, which are proportional to the actual widths of
		* the glyphs, and which have serifs. "MS Serif" is an example.
	     */
		FF_ROMAN = 0x01,
		/**
	     * Fonts with variable stroke widths, which are proportional to the actual widths of the
	     * glyphs, and which do not have serifs. "MS Sans Serif" is an example.
	     */
		FF_SWISS = 0x02,
		/**
	     * Fonts with constant stroke width, with or without serifs. Fixed-width fonts are
	     * usually modern. "Pica", "Elite", and "Courier New" are examples.
	     */
		FF_MODERN = 0x03,
		/**
	     * Fonts designed to look like handwriting. "Script" and "Cursive" are examples.
	     */
		FF_SCRIPT = 0x04,
		/**
	     * Novelty fonts. "Old English" is an example.
	     */
		FF_DECORATIVE = 0x05
		}
	}
