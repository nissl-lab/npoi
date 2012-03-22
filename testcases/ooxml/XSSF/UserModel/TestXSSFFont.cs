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

namespace NPOI.xssf.usermodel;

using NPOI.POIXMLException;
using NPOI.ss.usermodel.BaseTestFont;
using NPOI.ss.usermodel.Font;
using NPOI.ss.usermodel.FontCharset;
using NPOI.ss.usermodel.FontFamily;
using NPOI.ss.usermodel.FontScheme;
using NPOI.ss.usermodel.FontUnderline;
using NPOI.ss.usermodel.IndexedColors;
using NPOI.xssf.XSSFITestDataProvider;
using NPOI.xssf.XSSFTestDataSamples;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTBooleanProperty;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTFontName;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTFontScheme;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTFontSize;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTIntProperty;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTUnderlineProperty;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTVerticalAlignFontProperty;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STFontScheme;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STUnderlineValues;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STVerticalAlignRun;

public class TestXSSFFont : BaseTestFont{

	public TestXSSFFont() {
		base(XSSFITestDataProvider.instance);
	}

	public void TestDefaultFont() {
		baseTestDefaultFont("Calibri", (short) 220, IndexedColors.BLACK.GetIndex());
	}

	public void TestConstructor() {
		XSSFFont xssfFont=new XSSFFont();
		assertNotNull(xssfFont.GetCTFont());
	}

	public void TestBoldweight() {
		CTFont ctFont=CTFont.Factory.newInstance();
		CTBooleanProperty bool=ctFont.AddNewB();
		bool.SetVal(false);
		ctFont.SetBArray(0,bool);
		XSSFFont xssfFont=new XSSFFont(ctFont);
		Assert.AreEqual(false, xssfFont.GetBold());


		xssfFont.SetBold(true);
		Assert.AreEqual(ctFont.GetBList().Count,1);
		Assert.AreEqual(true, ctFont.GetBArray(0).GetVal());
	}

	public void TestCharSet() {
		CTFont ctFont=CTFont.Factory.newInstance();
		CTIntProperty prop=ctFont.AddNewCharset();
		prop.SetVal(FontCharset.ANSI.GetValue());

		ctFont.SetCharsetArray(0,prop);
		XSSFFont xssfFont=new XSSFFont(ctFont);
		Assert.AreEqual(Font.ANSI_CHARSET,xssfFont.GetCharSet());

		xssfFont.SetCharSet(FontCharset.DEFAULT);
		Assert.AreEqual(FontCharset.DEFAULT.GetValue(),ctFont.GetCharsetArray(0).GetVal());

		// Try with a few less usual ones:
		// Set with the Charset itself
      xssfFont.SetCharSet(FontCharset.RUSSIAN);
      Assert.AreEqual(FontCharset.RUSSIAN.GetValue(), xssfFont.GetCharSet());
      // And Set with the Charset index
      xssfFont.SetCharSet(FontCharset.ARABIC.GetValue());
      Assert.AreEqual(FontCharset.ARABIC.GetValue(), xssfFont.GetCharSet());
      
      // This one isn't allowed
      Assert.AreEqual(null, FontCharset.ValueOf(9999));
      try {
         xssfFont.SetCharSet(9999);
         Assert.Fail("Shouldn't be able to Set an invalid charset");
      } catch(POIXMLException e) {}
      
		
		// Now try with a few sample files
		
		// Normal charset
      XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("Formatting.xlsx");
      Assert.AreEqual(0, 
            workbook.GetSheetAt(0).GetRow(0).GetCell(0).GetCellStyle().GetFont().GetCharSet()
      );
		
		// GB2312 charact Set
      workbook = XSSFTestDataSamples.OpenSampleWorkbook("49273.xlsx");
      Assert.AreEqual(134, 
            workbook.GetSheetAt(0).GetRow(0).GetCell(0).GetCellStyle().GetFont().GetCharSet()
      );
	}

	public void TestFontName() {
		CTFont ctFont=CTFont.Factory.newInstance();
		CTFontName fname=ctFont.AddNewName();
		fname.SetVal("Arial");
		ctFont.SetNameArray(0,fname);

		XSSFFont xssfFont=new XSSFFont(ctFont);
		Assert.AreEqual("Arial", xssfFont.GetFontName());

		xssfFont.SetFontName("Courier");
		Assert.AreEqual("Courier",ctFont.GetNameArray(0).GetVal());
	}

	public void TestItalic() {
		CTFont ctFont=CTFont.Factory.newInstance();
		CTBooleanProperty bool=ctFont.AddNewI();
		bool.SetVal(false);
		ctFont.SetIArray(0,bool);

		XSSFFont xssfFont=new XSSFFont(ctFont);
		Assert.AreEqual(false, xssfFont.GetItalic());

		xssfFont.SetItalic(true);
		Assert.AreEqual(ctFont.GetIList().Count,1);
		Assert.AreEqual(true, ctFont.GetIArray(0).GetVal());
		Assert.AreEqual(true,ctFont.GetIArray(0).GetVal());
	}

	public void TestStrikeout() {
		CTFont ctFont=CTFont.Factory.newInstance();
		CTBooleanProperty bool=ctFont.AddNewStrike();
		bool.SetVal(false);
		ctFont.SetStrikeArray(0,bool);

		XSSFFont xssfFont=new XSSFFont(ctFont);
		Assert.AreEqual(false, xssfFont.GetStrikeout());

		xssfFont.SetStrikeout(true);
		Assert.AreEqual(ctFont.GetStrikeList().Count,1);
		Assert.AreEqual(true, ctFont.GetStrikeArray(0).GetVal());
		Assert.AreEqual(true,ctFont.GetStrikeArray(0).GetVal());
	}

	public void TestFontHeight() {
		CTFont ctFont=CTFont.Factory.newInstance();
		CTFontSize size=ctFont.AddNewSz();
		size.SetVal(11);
		ctFont.SetSzArray(0,size);

		XSSFFont xssfFont=new XSSFFont(ctFont);
		Assert.AreEqual(11,xssfFont.GetFontHeightInPoints());

		xssfFont.SetFontHeight(20);
		Assert.AreEqual(20.0, ctFont.GetSzArray(0).GetVal(), 0.0);
	}

	public void TestFontHeightInPoint() {
		CTFont ctFont=CTFont.Factory.newInstance();
		CTFontSize size=ctFont.AddNewSz();
		size.SetVal(14);
		ctFont.SetSzArray(0,size);

		XSSFFont xssfFont=new XSSFFont(ctFont);
		Assert.AreEqual(14,xssfFont.GetFontHeightInPoints());

		xssfFont.SetFontHeightInPoints((short)20);
		Assert.AreEqual(20.0, ctFont.GetSzArray(0).GetVal(), 0.0);
	}

	public void TestUnderline() {
		CTFont ctFont=CTFont.Factory.newInstance();
		CTUnderlineProperty underlinePropr=ctFont.AddNewU();
		underlinePropr.SetVal(STUnderlineValues.SINGLE);
		ctFont.SetUArray(0,underlinePropr);

		XSSFFont xssfFont=new XSSFFont(ctFont);
		Assert.AreEqual(Font.U_SINGLE, xssfFont.GetUnderline());

		xssfFont.SetUnderline(Font.U_DOUBLE);
		Assert.AreEqual(ctFont.GetUList().Count,1);
		Assert.AreEqual(STUnderlineValues.DOUBLE,ctFont.GetUArray(0).GetVal());

		xssfFont.SetUnderline(FontUnderline.DOUBLE_ACCOUNTING);
		Assert.AreEqual(ctFont.GetUList().Count,1);
		Assert.AreEqual(STUnderlineValues.DOUBLE_ACCOUNTING,ctFont.GetUArray(0).GetVal());
	}

	public void TestColor() {
		CTFont ctFont=CTFont.Factory.newInstance();
		CTColor color=ctFont.AddNewColor();
		color.SetIndexed(XSSFFont.DEFAULT_FONT_COLOR);
		ctFont.SetColorArray(0,color);

		XSSFFont xssfFont=new XSSFFont(ctFont);
		Assert.AreEqual(IndexedColors.BLACK.GetIndex(),xssfFont.GetColor());

		xssfFont.SetColor(IndexedColors.RED.GetIndex());
		Assert.AreEqual(IndexedColors.RED.GetIndex(), ctFont.GetColorArray(0).GetIndexed());
	}

	public void TestRgbColor() {
		CTFont ctFont=CTFont.Factory.newInstance();
		CTColor color=ctFont.AddNewColor();

		color.SetRgb(Int32.ToHexString(0xFFFFFF).GetBytes());
		ctFont.SetColorArray(0,color);

		XSSFFont xssfFont=new XSSFFont(ctFont);
		Assert.AreEqual(ctFont.GetColorArray(0).GetRgb()[0],xssfFont.GetXSSFColor().GetRgb()[0]);
		Assert.AreEqual(ctFont.GetColorArray(0).GetRgb()[1],xssfFont.GetXSSFColor().GetRgb()[1]);
		Assert.AreEqual(ctFont.GetColorArray(0).GetRgb()[2],xssfFont.GetXSSFColor().GetRgb()[2]);
		Assert.AreEqual(ctFont.GetColorArray(0).GetRgb()[3],xssfFont.GetXSSFColor().GetRgb()[3]);

		color.SetRgb(Int32.ToHexString(0xF1F1F1).GetBytes());
		XSSFColor newColor=new XSSFColor(color);
		xssfFont.SetColor(newColor);
		Assert.AreEqual(ctFont.GetColorArray(0).GetRgb()[2],newColor.GetRgb()[2]);
	}

	public void TestThemeColor() {
		CTFont ctFont=CTFont.Factory.newInstance();
		CTColor color=ctFont.AddNewColor();
		color.SetTheme(1);
		ctFont.SetColorArray(0,color);

		XSSFFont xssfFont=new XSSFFont(ctFont);
		Assert.AreEqual(ctFont.GetColorArray(0).GetTheme(),xssfFont.GetThemeColor());

		xssfFont.SetThemeColor(IndexedColors.RED.GetIndex());
		Assert.AreEqual(IndexedColors.RED.GetIndex(),ctFont.GetColorArray(0).GetTheme());
	}

	public void TestFamily() {
		CTFont ctFont=CTFont.Factory.newInstance();
		CTIntProperty family=ctFont.AddNewFamily();
		family.SetVal(FontFamily.MODERN.GetValue());
		ctFont.SetFamilyArray(0,family);

		XSSFFont xssfFont=new XSSFFont(ctFont);
		Assert.AreEqual(FontFamily.MODERN.GetValue(),xssfFont.GetFamily());
	}

	public void TestScheme() {
		CTFont ctFont=CTFont.Factory.newInstance();
		CTFontScheme scheme=ctFont.AddNewScheme();
		scheme.SetVal(STFontScheme.MAJOR);
		ctFont.SetSchemeArray(0,scheme);

		XSSFFont font=new XSSFFont(ctFont);
		Assert.AreEqual(FontScheme.MAJOR,font.GetScheme());

		font.SetScheme(FontScheme.NONE);
		Assert.AreEqual(STFontScheme.NONE,ctFont.GetSchemeArray(0).GetVal());
	}

	public void TestTypeOffset() {
		CTFont ctFont=CTFont.Factory.newInstance();
		CTVerticalAlignFontProperty valign=ctFont.AddNewVertAlign();
		valign.SetVal(STVerticalAlignRun.BASELINE);
		ctFont.SetVertAlignArray(0,valign);

		XSSFFont font=new XSSFFont(ctFont);
		Assert.AreEqual(Font.SS_NONE,font.GetTypeOffset());

		font.SetTypeOffset(XSSFFont.SS_SUPER);
		Assert.AreEqual(STVerticalAlignRun.SUPERSCRIPT,ctFont.GetVertAlignArray(0).GetVal());
	}
}

