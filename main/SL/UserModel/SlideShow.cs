using NPOI.Common.UserModel.Fonts;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.SL.UserModel
{
	public interface SlideShow<S,P> : ICloseable 
		where	S: Shape<S,P>
		where  P : TextParagraph<S,P, TextRun>
	{
		Slide<S, P> createSlide();

		List<Slide<S, P>> Slides { get; }

		/**
	     * Returns all slide masters.
	     * This doesn't include notes master and other arbitrary masters.
	     */
		List<MasterSheet<S, P>> SlideMasters { get; }

		/**
	     * Returns the current page size
	     *
	     * @return the page size
	     */
		Dimension PageSize { get; set; }

		/**
	     * Returns all Pictures of this slideshow.
	     * The returned {@link List} is unmodifiable.
	     * @return a {@link List} of {@link PictureData}.
	     */
		List<PictureData> PictureData { get; }

		/**
	     * Adds a picture to the presentation.
	     *
	     * @param pictureData       The bytes of the picture
	     * @param format            The format of the picture.
	     *
	     * @return the picture data reference.
	     */
		PictureData AddPicture(byte[] pictureData, PictureType format);

		/**
	     * Adds a picture to the presentation.
	     *
	     * @param is            The stream to read the image from
	     * @param format        The format of the picture.
	     *
	     * @return the picture data reference.
	     * @since 3.15 beta 1
	     */
	    PictureData AddPicture(InputStream _is, PictureType format);

		/**
	     * Adds a picture to the presentation.
	     *
	     * @param pict              The file containing the image to add
	     * @param format            The format of the picture.
	     *
	     * @return the picture data reference
	     * @since 3.15 beta 1
	     */
	    PictureData AddPicture(FileStream pict, PictureType format);

		/**
	     * check if a picture with this picture data already exists in this presentation
	     *
	     * @param pictureData The picture data to find in the SlideShow
	     * @return {@code null} if picture data is not found in this slideshow
	     * @since 3.15 beta 3
	     */
	    PictureData findPictureData(byte[] pictureData);

		/**
	     * Writes out the slideshow file the is represented by an instance of this
	     * class
	     *
	     * @param out
	     *            The OutputStream to write to.
	     * @throws IOException
	     *             If there is an unexpected IOException from the passed in
	     *             OutputStream
	     */
	    void write(OutputStream _out);

		/**
	     * @return an extractor for the slideshow metadata
	     *
	     * @since POI 4.0.0
	     */
	    POITextExtractor getMetadataTextExtractor();

		/**
	     * @return the instance which handles the persisting of the slideshow,
	     * which is either a subclass of {@link org.apache.poi.POIDocument}
	     * or {@link org.apache.poi.ooxml.POIXMLDocument}
	     *
	     * @since POI 4.0.0
	     */
	    Object getPersistDocument();

		/**
	     * Add an EOT font to the slideshow.
	     * An EOT or MTX font is a transformed True-Type (.ttf) or Open-Type (.otf) font.
	     * To transform a True-Type font use the sfntly library (see "see also" below)<p>
	     *
	     * (Older?) Powerpoint versions handle embedded fonts by converting them to .ttf files
	     * and put them into the Windows fonts directory. If the user is not allowed to install
	     * fonts, the slideshow can't be opened. While the slideshow is opened, its possible
	     * to copy the extracted .ttfs from the fonts directory. When the slideshow is closed,
	     * they will be removed.
	     *
	     * @param fontData the EOT font as stream
	     * @return the font info object containing the new font data
	     * @throws IOException if the fontData can't be saved or if the fontData is no EOT font
	     *
	     * @see <a href="http://www.w3.org/Submission/EOT">EOT specification</a>
	     * @see <a href="https://github.com/googlei18n/sfntly">googles sfntly library</a>
	     * @see <a href="https://github.com/kiwiwings/poi-font-mbender">Example on how to subset and embed fonts</a>
	     */
		FontInfo addFont(InputStream fontData);

		/**
	     * @return a list of registered fonts
	     */
	    List<FontInfo> getFonts();

		void Close();
	}
}