using System;

namespace NPOI.SL.UserModel
{
	public enum PlaceholderSize
	{
	        quarter, half, full
	}

	public interface PlaceholderDetails
	{
		Placeholder getPlaceholder();
	
	    /**
	     * Specifies that the corresponding shape should be represented by the generating application
	     * as a placeholder. When a shape is considered a placeholder by the generating application
	     * it can have special properties to alert the user that they may enter content into the shape.
	     * Different types of placeholders are allowed and can be specified by using the placeholder
	     * type attribute for this element
	     *
	     * @param placeholder The shape to use as placeholder or null if no placeholder should be set.
	     */
	    void setPlaceholder(Placeholder placeholder);
	
	    bool isVisible();
	
	    void setVisible(bool isVisible);
	
	    PlaceholderSize getSize();
	
	    void setSize(PlaceholderSize size);
	
	    /**
	     * If the placeholder shape or object stores text, this text is returned otherwise {@code null}.
	     *
	     * @return the text of the shape / placeholder
	     *
	     * @since POI 4.0.0
	     */
	    string getText();
	
	    /**
	     * If the placeholder shape or object stores text, the given text is stored otherwise this is a no-op.
	     *
	     * @param text the placeholder text
	     *
	     * @since POI 4.0.0
	     */
	    void setText(string text);


		/**
	     * @return the stored / fixed user specified date
	     *
	     * @since POI 5.2.0
	     */
		string getUserDate();

		/**
	     * @return Get the date format for the datetime placeholder
	     *
	     * @since POI 5.2.0
	     */
		DateTime getDateFormat();
	}
}