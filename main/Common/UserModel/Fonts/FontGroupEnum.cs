using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.Common.UserModel.Fonts
{
	/**
	 * Text runs can contain characters which will be handled (if configured) by a different font,
	 * because the default font (latin) doesn't contain corresponding glyphs.
	 *
	 * @since POI 3.17-beta2
	 *
	 * @see <a href="https://blogs.msdn.microsoft.com/officeinteroperability/2013/04/22/office-open-xml-themes-schemes-and-fonts/">Office Open XML Themes, Schemes, and Fonts</a>
	 */
	public enum FontGroupEnum
	{
		/** type for latin charset (default) - also used for unicode fonts like MS Arial Unicode */
	    LATIN,
	    /** type for east asian charsets - usually set as fallback for the latin font, e.g. something like MS Gothic or MS Mincho */
	    EAST_ASIAN,
	    /** type for symbol fonts */
	    SYMBOL,
	    /** type for complex scripts - see https://msdn.microsoft.com/en-us/library/windows/desktop/dd317698 */
	    COMPLEX_SCRIPT
	}
}
