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

namespace NPOI.SL.UserModel
{
	using NPOI.Common.UserModel.Fonts;
	using System;
	using System.Collections.Generic;
	using System.Drawing;
	using System.Text;

	/**
	     * Type of placeholder fields
	     */
	    enum FieldType
		{
	        SLIDE_NUMBER, DATE_TIME
	    }

	public interface TextRun
	{
		string GetRawText();

		void SetText(string text);

		TextCap TextCap { get; }

		/**
	     * Returns the font color.
	     * This usually returns a {@link SolidPaint}, but but also other classes are possible
	     *
	     * @return the font color/paint
	     *
	     * @see org.apache.poi.sl.draw.DrawPaint#getPaint(java.awt.Graphics2D, PaintStyle)
	     * @see SolidPaint#getSolidColor()
	     * @see org.apache.poi.sl.draw.DrawPaint#applyColorTransform(ColorStyle)
	     */
		IPaintStyle FontColor { get; set; }

		/**
	     * Sets the (solid) font color - convenience function
	     *
	     * @param color the color
	     */
		void SetFontColor(Color color);

		/**
	     * Returns the font size which is either set directly on this text run or
	     * given from the slide layout
	     *
	     * @return font size in points or null if font size is not set.
	     */
		/**
		* Sets the font size directly on this text run, if null is given, the
			* font size defaults to the values given from the slide layout
		*
		* @param fontSize font size in points, if null the underlying fontsize will be unset
		*/
		double FontSize { get; set; }

		/**
	     * Get the font family - convenience method for {@link #getFontInfo(FontGroup)}
	     *
	     * @return  font family or null if not set
	     */
		/**
		 * Specifies the typeface, or name of the font that is to be used for this text run -
		 * convenience method for calling {@link #setFontInfo(FontInfo, FontGroup)} with just a font name
		 *
		 * @param typeface  the font to apply to this text run.
		 *      The value of {@code null} removes the run specific font setting, so the default setting is activated again.
		 */
		string FontFamily { get; set; }

		/**
	     * Get the font family - convenience method for {@link #getFontInfo(FontGroup)}
	     *
	     * @param fontGroup the font group, i.e. the range of glpyhs to be covered.
	     *    if {@code null}, the font group matching the first character will be returned
	     *
	     * @return  font family or null if not set
	     */
		string GetFontFamily(FontGroupEnum fontGroup);

		/**
	     * Specifies the typeface, or name of the font that is to be used for this text run -
	     * convenience method for calling {@link #setFontInfo(FontInfo, FontGroup)} with just a font name
	     *
	     * @param typeface  the font to apply to this text run.
	     *      The value of {@code null} removes the run specific font setting, so the default setting is activated again.
	     * @param fontGroup the font group, i.e. the range of glpyhs to be covered.
	     *    if {@code null}, the font group matching the first character will be returned
	     */
		void SetFontFamily(String typeface, FontGroupEnum fontGroup);

		/**
	     * Get the font info for the given font group
	     *
	     * @param fontGroup the font group, i.e. the range of glpyhs to be covered.
	     *    if {@code null}, the font group matching the first character will be returned
	     * @return  font info or {@code null} if not set
	     *
	     * @since POI 3.17-beta2
	     */
		FontInfo getFontInfo(FontGroupEnum fontGroup);

		/**
	     * Specifies the font to be used for this text run.
	     *
	     * @param fontInfo the font to apply to this text run.
	     *      The value of {@code null} removes the run specific font setting, so the default setting is activated again.
	     * @param fontGroup the font group, i.e. the range of glpyhs to be covered. defaults to latin, if {@code null}.
	     *
	     * @since POI 3.17-beta2
	     */
		void setFontInfo(FontInfo fontInfo, FontGroupEnum fontGroup);

		/**
	     * @return true, if text is bold
	     */
		bool isBold();

		/**
	     * Sets the bold state
	     *
	     * @param bold set to true for bold text, false for normal weight
	     */
		void setBold(bool bold);

		/**
	     * @return true, if text is italic
	     */
		bool isItalic();

		/**
	     * Sets the italic state
	     *
	     * @param italic set to true for italic text, false for non-italics
	     */
		void setItalic(bool italic);

		/**
	     * @return true, if text is underlined
	     */
		bool isUnderlined();

		/**
	     * Sets the underlined state
	     *
	     * @param underlined set to true for underlined text, false for no underlining
	     */
		void setUnderlined(bool underlined);

		/**
	     * @return true, if text is stroked
	     */
		bool isStrikethrough();

		/**
	     * Sets the strikethrough state
	     *
	     * @param stroked set to true for stroked text, false for no stroking
	     */
		void setStrikethrough(bool stroked);

		/**
	     * @return true, if text is sub scripted
	     */
		bool isSubscript();

		/**
	     * @return true, if text is super scripted
	     */
		bool isSuperscript();

		/**
	     * @return the pitch and family id or -1 if not applicable
	     */
		byte getPitchAndFamily();

		/**
	     * Return the associated hyperlink
	     *
	     * @return the associated hyperlink or null if no hyperlink was set
	     *
	     * @since POI 3.14-Beta2
	     */
		Hyperlink<S,P> GetHyperlink<S, P>() where S : Shape<S, P> where P : TextParagraph<S, P, TextRun>;



		/**
	     * Creates a new hyperlink and assigns it to this text run.
	     * If the text run has already a hyperlink assigned, return it instead
	     *
	     * @return the associated hyperlink
	     *
	     * @since POI 3.14-Beta2
	     */
		Hyperlink<S,P> CreateHyperlink<S,P>() where S : Shape<S, P> where P : TextParagraph<S, P, TextRun>;

		/**
	     * Experimental method to determine the field type, e.g. slide number
	     *
	     * @return the field type or {@code null} if text run is not a field
	     */
		//@Internal
		FieldType getFieldType();

		/**
	     * @return the paragraph which contains this TextRun
	     *
	     * @since POI 4.1.0
	     */
		TextParagraph<S,P, TextRun> GetParagraph<S, P, T>() where S : Shape<S, P> where P : TextParagraph<S, P, TextRun>;
	}
}
