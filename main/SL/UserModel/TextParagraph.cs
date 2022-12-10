using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace NPOI.SL.UserModel
{
	/**
	     * Specifies a list of text alignment types
	     */
	public enum TextAlign
	{
		/**
		 * For horizontal text, left aligned.
		 * For vertical text, top aligned.
		 */
		LEFT,

		/**
		 * For horizontal text, centered.
		 * For vertical text, middle aligned.
		 */
		CENTER,

		/**
		 * For horizontal text, right aligned.
		 * For vertical text, bottom aligned.
		 */
		RIGHT,

		/**
		 * Align text so that it is justified across the whole line. It
		 * is smart in the sense that it will not justify sentences
		 * which are short
		 *
		 * For horizontal text, flush left and right.
		 * For vertical text, flush top and bottom.
		 */
		JUSTIFY,

		/**
		 * Kashida justify low.
		 */
		JUSTIFY_LOW,

		/**
		 * Distribute space between characters.
		 */
		DIST,

		/**
		 * Thai distribution justification.
		 */
		THAI_DIST
	}
	/**
	     *
	     */
	public enum FontAlign
	{
	        AUTO,
	
	        /**
	         * Characters hang from top of line height.
	         * Also known as "Hanging"
	         */
	        TOP,
	
	        /**
	         * Characters centered within line height.
	         */
	        CENTER,
	
	        /**
	         * Place characters on font baseline.
	         * Also known as "Roman"
	         */
	        BASELINE,
	
	        /**
	         * Characters are anchored to the very bottom of a single line.
	         * This is different than BASELINE because of letters such as "g", "q", and "y".
	         * Also known as "UpholdFixed"
	         */
	        BOTTOM
	    }

	public interface BulletStyle
	{
		String GetBulletCharacter();
		String GetBulletFont();

		/**
		 * The bullet point font size
		 * If bulletFontSize &gt;= 0, then space is a percentage of normal line height.
		 * If bulletFontSize &lt; 0, the absolute value in points
		 *
		 * @return the bullet point font size
		 */
		Double GetBulletFontSize();

		/**
		 * Convenience function to set a solid color
		 */
		void SetBulletFontColor(Color color);

		void SetBulletFontColor(PaintStyle color);

		/**
		 *
		 * @return the color of bullet characters within a given paragraph.
		 * A {@code null} value means to use the text font color.
		 */
		PaintStyle GetBulletFontColor();

		AutoNumberingScheme GetAutoNumberingScheme();
		/**
		 * Index (1-based) of the first auto number value, or null if auto numbering scheme
		 * wasn't assigned.
		 */
		int GetAutoNumberingStartAt();
	}

	public interface TextParagraph<S, P, T>: IEnumerable<T>
		where S : Shape<S,P>
		where P : TextParagraph<S, P, TextRun>
	{
		/**
	     * The amount of vertical white space before the paragraph
	     * This may be specified in two different ways, percentage spacing and font point spacing:
	     * <p>
	     * If spaceBefore &gt;= 0, then space is a percentage of normal line height.
	     * If spaceBefore &lt; 0, the absolute value in points
	     * </p>
	     *
	     * @return the vertical white space before the paragraph, or null if unset
	     */
	    double getSpaceBefore();

		/**
	     * Set the amount of vertical white space that will be present before the paragraph.
	     * This space is specified in either percentage or points:
	     * <p>
	     * If spaceBefore &gt;= 0, then space is a percentage of normal line height.
	     * If spaceBefore &lt; 0, the absolute value of linespacing is the spacing in points
	     * </p>
	     * Examples:
	     * <pre><code>
	     *      // The paragraph will be formatted to have a spacing before the paragraph text.
	     *      // The spacing will be 200% of the size of the largest text on each line
	     *      paragraph.setSpaceBefore(200);
	     *
	     *      // The spacing will be a size of 48 points
	     *      paragraph.setSpaceBefore(-48.0);
	     * </code></pre>
	     *
	     * @param spaceBefore the vertical white space before the paragraph, null to unset
	     */
	    void setSpaceBefore(double spaceBefore);

		/**
	     * The amount of vertical white space after the paragraph
	     * This may be specified in two different ways, percentage spacing and font point spacing:
	     * <p>
	     * If spaceBefore &gt;= 0, then space is a percentage of normal line height.
	     * If spaceBefore &lt; 0, the absolute value of linespacing is the spacing in points
	     * </p>
	     *
	     * @return the vertical white space after the paragraph or null, if unset
	     */
	    double getSpaceAfter();

		/**
	     * Set the amount of vertical white space that will be present after the paragraph.
	     * This space is specified in either percentage or points:
	     * <p>
	     * If spaceAfter &gt;= 0, then space is a percentage of normal line height.
	     * If spaceAfter &lt; 0, the absolute value of linespacing is the spacing in points
	     * </p>
	     * Examples:
	     * <pre><code>
	     *      // The paragraph will be formatted to have a spacing after the paragraph text.
	     *      // The spacing will be 200% of the size of the largest text on each line
	     *      paragraph.setSpaceAfter(200);
	     *
	     *      // The spacing will be a size of 48 points
	     *      paragraph.setSpaceAfter(-48.0);
	     * </code></pre>
	     *
	     * @param spaceAfter the vertical white space after the paragraph, null to unset
	     */
	    void setSpaceAfter(double spaceAfter);

		/**
	     * @return the left margin (in points) of the paragraph or null, if unset
	     */
	    double getLeftMargin();

		/**
	     * Specifies the left margin of the paragraph. This is specified in addition to the text body
	     * inset and applies only to this text paragraph. That is the text body Inset and the LeftMargin
	     * attributes are additive with respect to the text position.
	     *
	     * @param leftMargin the left margin (in points) or null to unset
	     */
	    void setLeftMargin(double leftMargin);

		/**
	     * Specifies the right margin of the paragraph. This is specified in addition to the text body
	     * inset and applies only to this text paragraph. That is the text body Inset and the RightMargin
	     * attributes are additive with respect to the text position.
	     *
	     * The right margin is not support and therefore ignored by the HSLF implementation.
	     *
	     * @return the right margin (in points) of the paragraph or null, if unset
	     */
	    double getRightMargin();

		/**
	     * @param rightMargin the right margin (in points) of the paragraph
	     */
	    void setRightMargin(double rightMargin);

		/**
	     * @return the indent (in points) applied to the first line of text in the paragraph.
	     *  or null, if unset
	     */
	    double getIndent();

		/**
	     * Specifies the indent size that will be applied to the first line of text in the paragraph.
	     *
	     * @param indent the indent (in points) applied to the first line of text in the paragraph
	     */
	    void setIndent(double indent);

		/**
	     * @return the text level of this paragraph (0-based). Default is 0.
	     */
	    int getIndentLevel();

		/**
	     * Specifies the particular level text properties that this paragraph will follow.
	     * The value for this attribute formats the text according to the corresponding level
	     * paragraph properties defined in the SlideMaster.
	     *
	     * @param level the level (0 ... 4)
	     */
	    void setIndentLevel(int level);

		/**
	     * Returns the vertical line spacing that is to be used within a paragraph.
	     * This may be specified in two different ways, percentage spacing and font point spacing:
	     * <p>
	     * If linespacing &gt;= 0, then linespacing is a percentage of normal line height.
	     * If linespacing &lt; 0, the absolute value of linespacing is the spacing in points
	     * </p>
	     *
	     * @return the vertical line spacing or null, if unset
	     */
	    double getLineSpacing();

		/**
	     * This element specifies the vertical line spacing that is to be used within a paragraph.
	     * This may be specified in two different ways, percentage spacing and font point spacing:
	     * <p>
	     * If linespacing &gt;= 0, then linespacing is a percentage of normal line height
	     * If linespacing &lt; 0, the absolute value of linespacing is the spacing in points
	     * </p>
	     * Examples:
	     * <pre><code>
	     *      // spacing will be 120% of the size of the largest text on each line
	     *      paragraph.setLineSpacing(120);
	     *
	     *      // spacing will be 200% of the size of the largest text on each line
	     *      paragraph.setLineSpacing(200);
	     *
	     *      // spacing will be 48 points
	     *      paragraph.setLineSpacing(-48.0);
	     * </code></pre>
	     *
	     * @param lineSpacing the vertical line spacing
	     */
	    void setLineSpacing(double lineSpacing);

		string getDefaultFontFamily();

		/**
	     * @return the default font size, in case its not set in the textrun or null, if unset
	     */
	    double getDefaultFontSize();

		/**
	     * Returns the alignment that is applied to the paragraph.
	     *
	     * If this attribute is omitted, then null is returned.
	     * User code can imply the value {@link org.apache.poi.sl.usermodel.TextParagraph.TextAlign#LEFT} then.
	     *
	     * @return alignment that is applied to the paragraph
	     */
	    TextAlign getTextAlign();

		/**
	     * Specifies the alignment that is to be applied to the paragraph.
	     * Possible values for this include left, right, centered, justified and distributed,
	     * see {@link org.apache.poi.sl.usermodel.TextParagraph.TextAlign}.
	     *
	     * @param align text align
	     */
	    void setTextAlign(TextAlign align);

		/**
	     * Returns the font alignment that is applied to the paragraph.
	     *
	     * If this attribute is omitted, then null is return,
	     * user code can imply the a value of {@link FontAlign#AUTO}
	     *
	     * @return alignment that is applied to the paragraph
	     */
	    FontAlign getFontAlign();

		/**
	     * @return the bullet style of the paragraph, if {@code null} then no bullets are used
	     */
	    BulletStyle getBulletStyle();

		/**
	     * Sets the bullet styles. If no styles are given, the bullets are omitted.
	     * Possible attributes are integer/double (bullet size), Color (bullet color),
	     * character (bullet character), string (bullet font), AutoNumberingScheme
	     */
	    void setBulletStyle(params object[] styles);

		/**
	     * @return the default size for a tab character within this paragraph in points, null if unset
	     */
	    double getDefaultTabSize();

		TextShape<S, P> ParentShape { get; }

		/**
	     * Fetch the text runs that are contained within this block of text
	     */
		List<T> getTextRuns();

		/**
	     * Convenience method to determine if this text paragraph is part of
	     * the slide header or footer
	     *
	     * @return true if this paragraph is part of a header or footer placeholder
	     *
	     * @since POI 3.15-beta2
	     */
	    bool isHeaderOrFooter();

		/**
	     * Get the {@link TabStop TabStops} - the list can't be and it's entries shouldn't be modified.
	     * Opposed to other properties, this method is not cascading to the master sheet,
	     * if the property is not defined on the normal slide level, i.e. the tabstops on
	     * different levels aren't merged.
	     *
	     * @return the tabstop collection or {@code null} if no tabstops are defined
	     *
	     * @since POI 4.0.0
	     */
	    //@SuppressWarnings("java:S1452")
	    List<TabStop> getTabStops();

		/**
	     * Set the {@link TabStop} collection
	     *
	     * @since POI 4.0.0
	     */
	    void addTabStops(double positionInPoints, TabStopTypeEnum tabStopType);

		/**
	     * Removes the tabstops of this paragraphs.
	     * This doesn't affect inherited tabstops, e.g. inherited by the slide master
	     *
	     * @since POI 4.0.0
	     */
	    void clearTabStops();
	}
}
