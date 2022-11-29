using NPOI.SS.Formula.Eval;
using System;
using System.Collections;
using System.Collections.Generic;

namespace NPOI.SL.UserModel
{
	/**
	     * Vertical Text Types
	     */
	    public enum TextDirection
	{
	        /**
	         * Horizontal text. This should be default.
	         */
	        HORIZONTAL,
	        /**
	         * Vertical orientation.
	         * (each line is 90 degrees rotated clockwise, so it goes
	         * from top to bottom; each next line is to the left from
	         * the previous one).
	         */
	        VERTICAL,
	        /**
	         * Vertical orientation.
	         * (each line is 270 degrees rotated clockwise, so it goes
	         * from bottom to top; each next line is to the right from
	         * the previous one).
	         * For HSLF: always interpreted by Powerpoint as HORIZONTAL.
	         */
	        VERTICAL_270,
	        /**
	         * Determines if all of the text is vertical
	         * ("one letter on top of another").
	         * For HSLF: not supported
	         */
	        STACKED
	}

	/**
	     * Specifies alist of auto-fit types.
	     * <p>
	     * Autofit specofies that a shape should be auto-fit to fully contain the text described within it.
	     * Auto-fitting is when text within a shape is scaled in order to contain all the text inside
	     * </p>
	     */
	    public enum TextAutofit
		{ 
	        /**
	         * Specifies that text within the text body should not be auto-fit to the bounding box.
	         * Auto-fitting is when text within a text box is scaled in order to remain inside
	         * the text box.
	         */
	        NONE,
	        /**
	         * Specifies that text within the text body should be normally auto-fit to the bounding box.
	         * Autofitting is when text within a text box is scaled in order to remain inside the text box.
	         *
	         * <p>
	         * <em>Example:</em> Consider the situation where a user is building a diagram and needs
	         * to have the text for each shape that they are using stay within the bounds of the shape.
	         * An easy way this might be done is by using NORMAL autofit
	         * </p>
	         */
	        NORMAL,
	        /**
	         * Specifies that a shape should be auto-fit to fully contain the text described within it.
	         * Auto-fitting is when text within a shape is scaled in order to contain all the text inside.
	         *
	         * <p>
	         * <em>Example:</em> Consider the situation where a user is building a diagram and needs to have
	         * the text for each shape that they are using stay within the bounds of the shape.
	         * An easy way this might be done is by using SHAPE autofit
	         * </p>
	         */
	        SHAPE
	}

	/**
	     * This enum represents a compromise for the handling of
	     * HSLF run types (see org.apache.poi.hslf.record.TextHeaderAtom) and
	     * XSLF placeholders (see org.apache.poi.xslf.usermodel.Placeholder).
	     * When a shape is considered a placeholder by the generating application
	     * it can have special properties to alert the user that they may enter content into the shape.
	     * 
	     * This enum and the handling around it may change significantly in future releases
	     */
	    public enum TextPlaceholderEnum
		{
	        /** Title placeholder shape text */
	        TITLE =0,
	        /** Body placeholder shape text */
	        BODY = 1,
	        /** Center title placeholder shape text */
	        CENTER_TITLE = 6,
	        /** Center body placeholder shape text */
	        CENTER_BODY = 5,
	        /** Half-sized body placeholder shape text */
	        HALF_BODY = 7,
	        /** Quarter-sized body placeholder shape text */
	        QUARTER_BODY = 8,
	        /** Notes placeholder shape text */
	        NOTES = 2,
	        /** Any other text */
	        OTHER = 4
	}

	public class TextPlaceholder
	{
		public int nativeId;
	
		public TextPlaceholder(int nativeId)
		{
			this.nativeId = nativeId;
		}

		public static TextPlaceholder fromNativeId(int nativeId)
		{
			foreach (TextPlaceholderEnum item in Enum.GetValues(typeof(TextPlaceholderEnum)))
			{
				if ((int)item == nativeId) return new TextPlaceholder(nativeId);
			}
			return null;
		}
	
		public static bool isTitle(int nativeId)
		{
			return (nativeId == (int)TextPlaceholderEnum.TITLE || nativeId == (int)TextPlaceholderEnum.CENTER_TITLE);
		}
	}

	public interface TextShape<S, P>: IEnumerable<P>, SimpleShape<S,P> 
		where S : Shape<S, P>
		where P : TextParagraph<S, P, TextRun>
	{
		/**
	     * Returns the text contained in this text frame, which has been made safe
	     * for printing and other use.
	     * 
	     * @return the text string for this textbox.
	     * 
	     * @since POI 3.14-Beta2
	     */
	    String getText();
	    
	    /**
	     * Sets (overwrites) the current text.
	     * Uses the properties of the first paragraph / textrun.
	     * Text paragraphs are split by \\r or \\n.
	     * New lines within text run are split by \\u000b
	     * 
	     * @param text the text string used by this object.
	     * 
	     * @return the last text run of the - potential split - text
	     */
	    TextRun setText(String text);
	
	    /**
	     * Adds the supplied text onto the end of the TextParagraphs,
	     * creating a new RichTextRun for it to sit in.
	     *
	     * @param text the text string to be appended.
	     * @param newParagraph if true, a new paragraph will be added,
	     *        which will contain the added text
	     *
	     * @since POI 3.14-Beta1
	     */
	    TextRun appendText(String text, bool newParagraph);
	    
	    /**
	     * @return the TextParagraphs for this text box
	     */
	    List<P> getTextParagraphs();
	
	    /**
	     * @return text shape margin
	     */
	    Insets2D getInsets();
	    
	    /**
	     * Sets the shape margins
	     *
	     * @param insets the new shape margins
	     */
	    void setInsets(Insets2D insets);
	
     /**
	     * Compute the cumulative height occupied by the text
	     * 
	     * @return the cumulative height occupied by the text
	     */
	    double getTextHeight();
	
	    /**
	     * Compute the cumulative height occupied by the text
	     * 
	     * @param graphics a customized graphics context, e.g. which contains font mappings
	     * 
	     * @return the cumulative height occupied by the text
	     * 
	     * @since POI 3.17-beta2
      */
	    //double getTextHeight(Graphics2D graphics);
	    
	    /**
	     * Returns the type of vertical alignment for the text.
	     *
	     * @return the type of vertical alignment
	     */
	    VerticalAlignment getVerticalAlignment();
	
	    /**
	     * Sets the type of vertical alignment for the text.
	     *
	     * @param vAlign - the type of alignment.
	     * A {@code null} values unsets this property.
	     */
	    void setVerticalAlignment(VerticalAlignment vAlign);
	    
	    /**
	     * Returns if the text is centered.
	     * If true and if the individual paragraph settings allow it,
	     * the whole text block will be displayed centered, i.e. its left and right
	     * margin will be maximized while still keeping the alignment of the paragraphs
	     *
	     * @return true, if the text anchor is horizontal centered
	     */
	    bool isHorizontalCentered();
	
	    /**
	     * Sets if the paragraphs are horizontal centered
	     *
	     * @param isCentered true, if the paragraphs are horizontal centered
	     * A {@code null} values unsets this property.
	     */
	    void setHorizontalCentered(bool isCentered);
	    
	    /**
	     * @return whether to wrap words within the bounding rectangle
	     */
	    bool getWordWrap();
	
	    /**
	     * @param wrap whether to wrap words within the bounding rectangle
	     */
	    void setWordWrap(bool wrap);
	    
	    /**
	     * @return vertical orientation of the text
	     */
	    TextDirection getTextDirection();
	
	    /**
	     * sets the vertical orientation
	     * @param orientation vertical orientation of the text
	     */
	    void setTextDirection(TextDirection orientation);
	
	    /**
	     * The text rotation can be independent specified from the shape rotation.
	     * For XSLF this can be an arbitrary degree, for HSLF the degree is given in steps of 90 degrees
	     * 
	     * @return text rotation in degrees, returns null if no rotation is given
	     */
	    double getTextRotation();
	    
	    /**
	     * Sets the text rotation.
	     * For XSLF this can ben an arbitrary degree, for HSLF the rotation is rounded to next 90 degree step
	     * 
	     * @param rotation the text rotation, or null to unset the rotation
	     */
	    void setTextRotation(Double rotation);
	    
	    /**
	     * Sets the text placeholder
	     */
		void setTextPlaceholder(TextPlaceholder placeholder);
	    
	    /**
	     * @return the text placeholder
	     */
	    TextPlaceholder getTextPlaceholder();
	    
	    /**
	     * Adjust the size of the shape so it encompasses the text inside it.
	     *
	     * @return a {@code Rectangle2D} that is the bounds of this shape.
	     * 
	     * @since POI 3.17-beta2
	     */
	//    Rectangle2D resizeToFitText();
	
	    /**
	     * Adjust the size of the shape so it encompasses the text inside it.
	     *
	     * @param graphics a customized graphics context, e.g. which contains font mappings
	     *
	     * @return a {@code Rectangle2D} that is the bounds of this shape.
	     * 
	     * @since POI 3.17-beta2
	     */
//	    Rectangle2D resizeToFitText(Graphics2D graphics);
	}
}