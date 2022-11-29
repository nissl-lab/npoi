namespace NPOI.SL.UserModel
{
	/**
	 * Specifies a list of available anchoring types for text
	 *
	 * <!-- FIXME: Identical to {@link org.apache.poi.ss.usermodel.VerticalAlignment}. Should merge these to
	 * {@link org.apache.poi.common.usermodel}.VerticalAlignment in the future. -->
	 */
	public enum VerticalAlignment
	{
	    /**
	     * Anchor the text at the top of the bounding rectangle
	     */
	    TOP,
	
	    /**
	     * Anchor the text at the middle of the bounding rectangle
	     */
	    MIDDLE,
	
	    /**
	     * Anchor the text at the bottom of the bounding rectangle.
	     */
	    BOTTOM,
	
	    /**
	     * Anchor the text so that it is justified vertically.
	     * <p>
	     * When text is horizontal, this spaces out the actual lines of
	     * text and is almost always identical in behavior to
	     * {@link #DISTRIBUTED} (special case: if only 1 line, then anchored at top).
	     * </p>
	     * <p>
	     * When text is vertical, then it justifies the letters
	     * vertically. This is different than {@link #DISTRIBUTED},
	     * because in some cases such as very little text in a line,
	     * it will not justify.
     * </p>
	     */
	    JUSTIFIED,
	
	    /**
	     * Anchor the text so that it is distributed vertically.
	     * <p>
	     * When text is horizontal, this spaces out the actual lines
	     * of text and is almost always identical in behavior to
	     * {@link #JUSTIFIED} (special case: if only 1 line, then anchored in middle).
	     * </p>
	     * <p>
	     * When text is vertical, then it distributes the letters vertically.
	     * This is different than {@link #JUSTIFIED}, because it always forces distribution
	     * of the words, even if there are only one or two words in a line.
	     */
	    DISTRIBUTED
	}
}