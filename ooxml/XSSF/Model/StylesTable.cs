/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI.XSSF.model;

using java.io.IOException;
using java.io.Stream;
using java.io.Stream;
using java.util.*;
using java.util.Map.Entry;

using NPOI.SS.usermodel.FontFamily;
using NPOI.SS.usermodel.FontScheme;
using NPOI.SS.usermodel.BuiltinFormats;
using NPOI.XSSF.usermodel.XSSFCellStyle;
using NPOI.XSSF.usermodel.XSSFFont;
using NPOI.XSSF.usermodel.extensions.XSSFCellBorder;
using NPOI.XSSF.usermodel.extensions.XSSFCellFill;
using org.apache.poi.POIXMLDocumentPart;
using org.apache.xmlbeans.XmlException;
using org.apache.xmlbeans.XmlOptions;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTBorders;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTCellStyleXfs;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTCellXfs;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTDxfs;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTFill;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTFills;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTFonts;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTNumFmt;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTNumFmts;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTStylesheet;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTXf;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.StyleSheetDocument;
using org.apache.poi.Openxml4j.opc.PackagePart;
using org.apache.poi.Openxml4j.opc.PackageRelationship;


/**
 * Table of styles shared across all sheets in a workbook.
 *
 * @author ugo
 */
public class StylesTable : POIXMLDocumentPart {
	private Dictionary<int, String> numberFormats = new LinkedDictionary<int,String>();
	private List<XSSFFont> fonts = new ArrayList<XSSFFont>();
	private List<XSSFCellFill> Fills = new ArrayList<XSSFCellFill>();
	private List<XSSFCellBorder> borders = new ArrayList<XSSFCellBorder>();
	private List<CTXf> styleXfs = new ArrayList<CTXf>();
	private List<CTXf> xfs = new ArrayList<CTXf>();

	private List<CTDxf> dxfs = new ArrayList<CTDxf>();

	/**
	 * The first style id available for use as a custom style
	 */
	public static int FIRST_CUSTOM_STYLE_ID = BuiltinFormats.FIRST_USER_DEFINED_FORMAT_INDEX + 1;

	private StyleSheetDocument doc;
	private ThemesTable theme;

	/**
	 * Create a new, empty StylesTable
	 */
	public StylesTable() {
		base();
		doc = StyleSheetDocument.Factory.newInstance();
		doc.AddNewStyleSheet();
		// Initialization required in order to make the document Readable by MSExcel
		Initialize();
	}

	public StylesTable(PackagePart part, PackageRelationship rel)  {
		base(part, rel);
		ReadFrom(part.GetStream());
	}

	public ThemesTable GetTheme() {
        return theme;
    }

    public void SetTheme(ThemesTable theme) {
        this.theme = theme;
        
        // Pass the themes table along to things which need to 
        //  know about it, but have already been Created by now
        for(XSSFFont font : fonts) {
           font.SetThemesTable(theme);
        }
        for(XSSFCellBorder border : borders) {
           border.SetThemesTable(theme);
        }
    }

	/**
	 * Read this shared styles table from an XML file.
	 *
	 * @param is The input stream Containing the XML document.
	 * @throws IOException if an error occurs while Reading.
	 */
    @SuppressWarnings("deprecation") //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
	protected void ReadFrom(Stream is)  {
		try {
			doc = StyleSheetDocument.Factory.Parse(is);

            CTStylesheet styleSheet = doc.GetStyleSheet();

            // Grab all the different bits we care about
			CTNumFmts ctfmts = styleSheet.GetNumFmts();
            if( ctfmts != null){
                for (CTNumFmt nfmt : ctfmts.GetNumFmtArray()) {
                    numberFormats.Put((int)nfmt.GetNumFmtId(), nfmt.GetFormatCode());
                }
            }

            CTFonts ctfonts = styleSheet.GetFonts();
            if(ctfonts != null){
				int idx = 0;
				for (CTFont font : ctfonts.GetFontArray()) {
				   // Create the font and save it. Themes Table supplied later
					XSSFFont f = new XSSFFont(font, idx);
					fonts.Add(f);
					idx++;
				}
			}
            CTFills ctFills = styleSheet.GetFills();
            if(ctFills != null){
                for (CTFill fill : ctFills.GetFillArray()) {
                    Fills.Add(new XSSFCellFill(Fill));
                }
            }

            CTBorders ctborders = styleSheet.GetBorders();
            if(ctborders != null) {
                for (CTBorder border : ctborders.GetBorderArray()) {
                    borders.Add(new XSSFCellBorder(border));
                }
            }

            CTCellXfs cellXfs = styleSheet.GetCellXfs();
            if(cellXfs != null) xfs.AddAll(Arrays.asList(cellXfs.GetXfArray()));

            CTCellStyleXfs cellStyleXfs = styleSheet.GetCellStyleXfs();
            if(cellStyleXfs != null) styleXfs.AddAll(Arrays.asList(cellStyleXfs.GetXfArray()));

            CTDxfs styleDxfs = styleSheet.GetDxfs();
			if(styleDxfs != null) dxfs.AddAll(Arrays.asList(styleDxfs.GetDxfArray()));

		} catch (XmlException e) {
			throw new IOException(e.GetLocalizedMessage());
		}
	}

	// ===========================================================
	//  Start of style related Getters and Setters
	// ===========================================================

	public String GetNumberFormatAt(int idx) {
		return numberFormats.Get(idx);
	}

	public int PutNumberFormat(String fmt) {
		if (numberFormats.ContainsValue(fmt)) {
			// Find the key, and return that
			for(int key : numberFormats.keySet() ) {
				if(numberFormats.Get(key).Equals(fmt)) {
					return key;
				}
			}
			throw new InvalidOperationException("Found the format, but couldn't figure out where - should never happen!");
		}

		// Find a spare key, and add that
		int newKey = FIRST_CUSTOM_STYLE_ID;
		while(numberFormats.ContainsKey(newKey)) {
			newKey++;
		}
		numberFormats.Put(newKey, fmt);
		return newKey;
	}

	public XSSFFont GetFontAt(int idx) {
		return fonts.Get(idx);
	}

	/**
	 * Records the given font in the font table.
	 * Will re-use an existing font index if this
	 *  font matches another, EXCEPT if forced
	 *  registration is requested.
	 * This allows people to create several fonts
	 *  then customise them later.
	 * Note - End Users probably want to call
	 *  {@link XSSFFont#registerTo(StylesTable)}
	 */
	public int PutFont(XSSFFont font, bool forceRegistration) {
		int idx = -1;
		if(!forceRegistration) {
			idx = fonts.IndexOf(font);
		}

		if (idx != -1) {
			return idx;
		}
		
		idx = fonts.Count;
		fonts.Add(font);
		return idx;
	}
	public int PutFont(XSSFFont font) {
		return PutFont(font, false);
	}

	public XSSFCellStyle GetStyleAt(int idx) {
		int styleXfId = 0;

		// 0 is the empty default
		if(xfs.Get(idx).GetXfId() > 0) {
			styleXfId = (int) xfs.Get(idx).GetXfId();
		}

		return new XSSFCellStyle(idx, styleXfId, this, theme);
	}
	public int PutStyle(XSSFCellStyle style) {
		CTXf mainXF = style.GetCoreXf();

		if(! xfs.Contains(mainXF)) {
			xfs.Add(mainXF);
		}
		return xfs.IndexOf(mainXF);
	}

	public XSSFCellBorder GetBorderAt(int idx) {
		return borders.Get(idx);
	}

	public int PutBorder(XSSFCellBorder border) {
		int idx = borders.IndexOf(border);
		if (idx != -1) {
			return idx;
		}
		borders.Add(border);
		border.SetThemesTable(theme);
		return borders.Count - 1;
	}

	public XSSFCellFill GetFillAt(int idx) {
		return Fills.Get(idx);
	}

	public List<XSSFCellBorder> GetBorders(){
		return borders;
	}

	public List<XSSFCellFill> GetFills(){
		return Fills;
	}

	public List<XSSFFont> GetFonts(){
		return fonts;
	}

	public Dictionary<int, String> GetNumberFormats(){
		return numberFormats;
	}

	public int PutFill(XSSFCellFill Fill) {
		int idx = Fills.IndexOf(Fill);
		if (idx != -1) {
			return idx;
		}
		Fills.Add(Fill);
		return Fills.Count - 1;
	}

	public CTXf GetCellXfAt(int idx) {
		return xfs.Get(idx);
	}
	public int PutCellXf(CTXf cellXf) {
		xfs.Add(cellXf);
		return xfs.Count;
	}
   public void ReplaceCellXfAt(int idx, CTXf cellXf) {
      xfs.Set(idx, cellXf);
   }

	public CTXf GetCellStyleXfAt(int idx) {
		return styleXfs.Get(idx);
	}
	public int PutCellStyleXf(CTXf cellStyleXf) {
		styleXfs.Add(cellStyleXf);
		return styleXfs.Count;
	}
	public void ReplaceCellStyleXfAt(int idx, CTXf cellStyleXf) {
	   styleXfs.Set(idx, cellStyleXf);
	}
	
	/**
	 * get the size of cell styles
	 */
	public int GetNumCellStyles(){
        // Each cell style has a unique xfs entry
        // Several might share the same styleXfs entry
        return xfs.Count;
	}

	/**
	 * For unit testing only
	 */
	public int _GetNumberFormatSize() {
		return numberFormats.Count;
	}

	/**
	 * For unit testing only
	 */
	public int _GetXfsSize() {
		return xfs.Count;
	}
	/**
	 * For unit testing only
	 */
	public int _GetStyleXfsSize() {
		return styleXfs.Count;
	}
	/**
	 * For unit testing only!
	 */
	public CTStylesheet GetCTStylesheet() {
		return doc.GetStyleSheet();
	}
    public int _GetDXfsSize() {
        return dxfs.Count;
    }


	/**
	 * Write this table out as XML.
	 *
	 * @param out The stream to write to.
	 * @throws IOException if an error occurs while writing.
	 */
	public void WriteTo(Stream out)  {
		XmlOptions options = new XmlOptions(DEFAULT_XML_OPTIONS);

		// Work on the current one
		// Need to do this, as we don't handle
		//  all the possible entries yet
        CTStylesheet styleSheet = doc.GetStyleSheet();

		// Formats
		CTNumFmts formats = CTNumFmts.Factory.newInstance();
		formats.SetCount(numberFormats.Count);
		for (Entry<int, String> fmt : numberFormats.entrySet()) {
			CTNumFmt ctFmt = formats.AddNewNumFmt();
			ctFmt.SetNumFmtId(fmt.GetKey());
			ctFmt.SetFormatCode(fmt.GetValue());
		}
		styleSheet.SetNumFmts(formats);

		int idx;
		// Fonts
		CTFonts ctFonts = CTFonts.Factory.newInstance();
		ctFonts.SetCount(fonts.Count);
		CTFont[] ctfnt = new CTFont[fonts.Count];
		idx = 0;
		for(XSSFFont f : fonts) ctfnt[idx++] = f.GetCTFont();
		ctFonts.SetFontArray(ctfnt);
		styleSheet.SetFonts(ctFonts);

		// Fills
		CTFills ctFills = CTFills.Factory.newInstance();
		ctFills.SetCount(Fills.Count);
		CTFill[] ctf = new CTFill[Fills.Count];
		idx = 0;
		for(XSSFCellFill f : Fills) ctf[idx++] = f.GetCTFill();
		ctFills.SetFillArray(ctf);
		styleSheet.SetFills(ctFills);

		// Borders
		CTBorders ctBorders = CTBorders.Factory.newInstance();
		ctBorders.SetCount(borders.Count);
		CTBorder[] ctb = new CTBorder[borders.Count];
		idx = 0;
		for(XSSFCellBorder b : borders) ctb[idx++] = b.GetCTBorder();
		ctBorders.SetBorderArray(ctb);
		styleSheet.SetBorders(ctBorders);

		// Xfs
		if(xfs.Count > 0) {
			CTCellXfs ctXfs = CTCellXfs.Factory.newInstance();
			ctXfs.SetCount(xfs.Count);
			ctXfs.SetXfArray(
					xfs.ToArray(new CTXf[xfs.Count])
			);
			styleSheet.SetCellXfs(ctXfs);
		}

		// Style xfs
		if(styleXfs.Count > 0) {
			CTCellStyleXfs ctSXfs = CTCellStyleXfs.Factory.newInstance();
			ctSXfs.SetCount(styleXfs.Count);
			ctSXfs.SetXfArray(
					styleXfs.ToArray(new CTXf[styleXfs.Count])
			);
			styleSheet.SetCellStyleXfs(ctSXfs);
		}

		// Style dxfs
		if(dxfs.Count > 0) {
			CTDxfs ctDxfs = CTDxfs.Factory.newInstance();
			ctDxfs.SetCount(dxfs.Count);
			ctDxfs.SetDxfArray(dxfs.ToArray(new CTDxf[dxfs.Count])
			);
			styleSheet.SetDxfs(ctDxfs);
		}

		// Save
		doc.save(out, options);
	}

	
	protected void Commit()  {
		PackagePart part = GetPackagePart();
		Stream out = part.GetStream();
		WriteTo(out);
		out.Close();
	}

	private void Initialize() {
		//CTFont ctFont = CreateDefaultFont();
		XSSFFont xssfFont = CreateDefaultFont();
		fonts.Add(xssfFont);

		CTFill[] ctFill = CreateDefaultFills();
		Fills.Add(new XSSFCellFill(ctFill[0]));
		Fills.Add(new XSSFCellFill(ctFill[1]));

		CTBorder ctBorder = CreateDefaultBorder();
		borders.Add(new XSSFCellBorder(ctBorder));

		CTXf styleXf = CreateDefaultXf();
		styleXfs.Add(styleXf);
		CTXf xf = CreateDefaultXf();
		xf.SetXfId(0);
		xfs.Add(xf);
	}

	private static CTXf CreateDefaultXf() {
		CTXf ctXf = CTXf.Factory.newInstance();
		ctXf.SetNumFmtId(0);
		ctXf.SetFontId(0);
		ctXf.SetFillId(0);
		ctXf.SetBorderId(0);
		return ctXf;
	}
	private static CTBorder CreateDefaultBorder() {
		CTBorder ctBorder = CTBorder.Factory.newInstance();
		ctBorder.AddNewBottom();
		ctBorder.AddNewTop();
		ctBorder.AddNewLeft();
		ctBorder.AddNewRight();
		ctBorder.AddNewDiagonal();
		return ctBorder;
	}


	private static CTFill[] CreateDefaultFills() {
		CTFill[] ctFill = new CTFill[]{CTFill.Factory.newInstance(),CTFill.Factory.newInstance()};
		ctFill[0].AddNewPatternFill().SetPatternType(STPatternType.NONE);
		ctFill[1].AddNewPatternFill().SetPatternType(STPatternType.DARK_GRAY);
		return ctFill;
	}

	private static XSSFFont CreateDefaultFont() {
		CTFont ctFont = CTFont.Factory.newInstance();
		XSSFFont xssfFont=new XSSFFont(ctFont, 0);
		xssfFont.SetFontHeightInPoints(XSSFFont.DEFAULT_FONT_SIZE);
		xssfFont.SetColor(XSSFFont.DEFAULT_FONT_COLOR);//SetTheme
		xssfFont.SetFontName(XSSFFont.DEFAULT_FONT_NAME);
		xssfFont.SetFamily(FontFamily.SWISS);
		xssfFont.SetScheme(FontScheme.MINOR);
		return xssfFont;
	}

	public CTDxf GetDxfAt(int idx) {
		return dxfs.Get(idx);
	}

	public int PutDxf(CTDxf dxf) {
		this.dxfs.Add(dxf);
		return this.dxfs.Count;
	}

	public XSSFCellStyle CreateCellStyle() {
		CTXf xf = CTXf.Factory.newInstance();
		xf.SetNumFmtId(0);
		xf.SetFontId(0);
		xf.SetFillId(0);
		xf.SetBorderId(0);
		xf.SetXfId(0);
		int xfSize = styleXfs.Count;
		int indexXf = PutCellXf(xf);
		return new XSSFCellStyle(indexXf - 1, xfSize - 1, this, theme);
	}

	/**
	 * Finds a font that matches the one with the supplied attributes
	 */
	public XSSFFont FindFont(short boldWeight, short color, short fontHeight, String name, bool italic, bool strikeout, short typeOffSet, byte underline) {
		for (XSSFFont font : fonts) {
			if (	(font.GetBoldweight() == boldWeight)
					&& font.GetColor() == color
					&& font.GetFontHeight() == fontHeight
					&& font.GetFontName().Equals(name)
					&& font.GetItalic() == italic
					&& font.GetStrikeout() == strikeout
					&& font.GetTypeOffSet() == typeOffset
					&& font.GetUnderline() == underline)
			{
				return font;
			}
		}
		return null;
	}
}






