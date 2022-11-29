using System;

namespace NPOI.SL.UserModel
{
	public interface Sheet<S, P> : ShapeContainer<S,P>
		where S : Shape<S,P>
		where P : TextParagraph<S,P,TextRun>
	{
		SlideShow<S, P> getSlideShow();

		/**
	     * @return whether shapes on the master sheet should be shown. By default master graphics is turned off.
	     * Sheets that support the notion of master (slide, slideLayout) should override it and
	     * check this setting in the sheet XML
	     */
	    bool getFollowMasterGraphics();

		MasterSheet<S, P> getMasterSheet();

		Background<S, P> getBackground();

		/**
	     * Convenience method to draw a sheet to a graphics context
	     */
	    //void draw(Graphics2D graphics);

		/**
	     * Get the placeholder details for the given placeholder type. Not all placeholders are also shapes -
	     * this is especially true for old HSLF slideshows, which notes have header/footers elements which
	     * aren't shapes.
	     *
	     * @param placeholder the placeholder type
	     * @return the placeholder details or {@code null}, if the placeholder isn't contained in the sheet
	     *
	     * @since POI 4.0.0
	     */
	    PlaceholderDetails getPlaceholderDetails(Placeholder placeholder);
	}
}