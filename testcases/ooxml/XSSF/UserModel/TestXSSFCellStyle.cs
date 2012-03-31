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

namespace NPOI.XSSF.UserModel;

using junit.framework.TestCase;

using NPOI.hssf.UserModel.HSSFCellStyle;
using NPOI.hssf.UserModel.HSSFWorkbook;
using NPOI.ss.UserModel.CellStyle;
using NPOI.ss.UserModel.HorizontalAlignment;
using NPOI.ss.UserModel.IndexedColors;
using NPOI.ss.UserModel.VerticalAlignment;
using NPOI.XSSF.XSSFTestDataSamples;
using NPOI.XSSF.Model.StylesTable;
using NPOI.XSSF.UserModel.extensions.XSSFCellBorder;
using NPOI.XSSF.UserModel.extensions.XSSFCellFill;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTCellXfs;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTFill;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTStylesheet;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTXf;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STHorizontalAlignment;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STVerticalAlignment;


public class TestXSSFCellStyle  {

	private StylesTable stylesTable;
	private CTBorder ctBorderA;
	private CTFill ctFill;
	private CTFont ctFont;
	private CTXf cellStyleXf;
	private CTXf cellXf;
	private CTCellXfs cellXfs;
	private XSSFCellStyle cellStyle;
	private CTStylesheet ctStylesheet;

	
	protected void SetUp() {
		stylesTable = new StylesTable();

		ctStylesheet = stylesTable.GetCTStylesheet();

		ctBorderA = CTBorder.Factory.newInstance();
		XSSFCellBorder borderA = new XSSFCellBorder(ctBorderA);
		long borderId = stylesTable.PutBorder(borderA);
		Assert.AreEqual(1, borderId);

		XSSFCellBorder borderB = new XSSFCellBorder();
		Assert.AreEqual(1, stylesTable.PutBorder(borderB));

		ctFill = CTFill.Factory.newInstance();
		XSSFCellFill fill = new XSSFCellFill(ctFill);
		long FillId = stylesTable.PutFill(Fill);
		Assert.AreEqual(2, FillId);

		ctFont = CTFont.Factory.newInstance();
		XSSFFont font = new XSSFFont(ctFont);
		long fontId = stylesTable.PutFont(font);
		Assert.AreEqual(1, fontId);

		cellStyleXf = ctStylesheet.AddNewCellStyleXfs().AddNewXf();
		cellStyleXf.SetBorderId(1);
		cellStyleXf.SetFillId(1);
		cellStyleXf.SetFontId(1);

		cellXfs = ctStylesheet.AddNewCellXfs();
		cellXf = cellXfs.AddNewXf();
		cellXf.SetXfId(1);
		cellXf.SetBorderId(1);
		cellXf.SetFillId(1);
		cellXf.SetFontId(1);
		stylesTable.PutCellStyleXf(cellStyleXf);
		stylesTable.PutCellXf(cellXf);
		cellStyle = new XSSFCellStyle(1, 1, stylesTable, null);
	}

	public void TestGetSetBorderBottom() {
        //default values
        Assert.AreEqual(CellStyle.BORDER_NONE, cellStyle.GetBorderBottom());

        int num = stylesTable.GetBorders().Count;
        cellStyle.SetBorderBottom(CellStyle.BORDER_MEDIUM);
        Assert.AreEqual(CellStyle.BORDER_MEDIUM, cellStyle.GetBorderBottom());
        //a new border has been Added
        Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);
        //id of the Created border
        int borderId = (int)cellStyle.GetCoreXf().GetBorderId();
        Assert.IsTrue(borderId > 0);
        //check Changes in the underlying xml bean
        CTBorder ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
        Assert.AreEqual(STBorderStyle.MEDIUM, ctBorder.GetBottom().GetStyle());

        num = stylesTable.GetBorders().Count;
        //setting the same border multiple times should not change borderId
        for (int i = 0; i < 3; i++) {
            cellStyle.SetBorderBottom(CellStyle.BORDER_MEDIUM);
            Assert.AreEqual(CellStyle.BORDER_MEDIUM, cellStyle.GetBorderBottom());
        }
        Assert.AreEqual(borderId, cellStyle.GetCoreXf().GetBorderId());
        Assert.AreEqual(num, stylesTable.GetBorders().Count);
        Assert.AreSame(ctBorder, stylesTable.GetBorderAt(borderId).GetCTBorder());

        //setting border to none Removes the <bottom> element
        cellStyle.SetBorderBottom(CellStyle.BORDER_NONE);
        Assert.AreEqual(num, stylesTable.GetBorders().Count);
        borderId = (int)cellStyle.GetCoreXf().GetBorderId();
        ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
        Assert.IsFalse(ctBorder.IsSetBottom());
    }

    public void TestGetSetBorderRight() {
        //default values
        Assert.AreEqual(CellStyle.BORDER_NONE, cellStyle.GetBorderRight());

        int num = stylesTable.GetBorders().Count;
        cellStyle.SetBorderRight(CellStyle.BORDER_MEDIUM);
        Assert.AreEqual(CellStyle.BORDER_MEDIUM, cellStyle.GetBorderRight());
        //a new border has been Added
        Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);
        //id of the Created border
        int borderId = (int)cellStyle.GetCoreXf().GetBorderId();
        Assert.IsTrue(borderId > 0);
        //check Changes in the underlying xml bean
        CTBorder ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
        Assert.AreEqual(STBorderStyle.MEDIUM, ctBorder.GetRight().GetStyle());

        num = stylesTable.GetBorders().Count;
        //setting the same border multiple times should not change borderId
        for (int i = 0; i < 3; i++) {
            cellStyle.SetBorderRight(CellStyle.BORDER_MEDIUM);
            Assert.AreEqual(CellStyle.BORDER_MEDIUM, cellStyle.GetBorderRight());
        }
        Assert.AreEqual(borderId, cellStyle.GetCoreXf().GetBorderId());
        Assert.AreEqual(num, stylesTable.GetBorders().Count);
        Assert.AreSame(ctBorder, stylesTable.GetBorderAt(borderId).GetCTBorder());

        //setting border to none Removes the <right> element
        cellStyle.SetBorderRight(CellStyle.BORDER_NONE);
        Assert.AreEqual(num, stylesTable.GetBorders().Count);
        borderId = (int)cellStyle.GetCoreXf().GetBorderId();
        ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
        Assert.IsFalse(ctBorder.IsSetRight());
    }

	public void TestGetSetBorderLeft() {
        //default values
        Assert.AreEqual(CellStyle.BORDER_NONE, cellStyle.GetBorderLeft());

        int num = stylesTable.GetBorders().Count;
        cellStyle.SetBorderLeft(CellStyle.BORDER_MEDIUM);
        Assert.AreEqual(CellStyle.BORDER_MEDIUM, cellStyle.GetBorderLeft());
        //a new border has been Added
        Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);
        //id of the Created border
        int borderId = (int)cellStyle.GetCoreXf().GetBorderId();
        Assert.IsTrue(borderId > 0);
        //check Changes in the underlying xml bean
        CTBorder ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
        Assert.AreEqual(STBorderStyle.MEDIUM, ctBorder.GetLeft().GetStyle());

        num = stylesTable.GetBorders().Count;
        //setting the same border multiple times should not change borderId
        for (int i = 0; i < 3; i++) {
            cellStyle.SetBorderLeft(CellStyle.BORDER_MEDIUM);
            Assert.AreEqual(CellStyle.BORDER_MEDIUM, cellStyle.GetBorderLeft());
        }
        Assert.AreEqual(borderId, cellStyle.GetCoreXf().GetBorderId());
        Assert.AreEqual(num, stylesTable.GetBorders().Count);
        Assert.AreSame(ctBorder, stylesTable.GetBorderAt(borderId).GetCTBorder());

        //setting border to none Removes the <left> element
        cellStyle.SetBorderLeft(CellStyle.BORDER_NONE);
        Assert.AreEqual(num, stylesTable.GetBorders().Count);
        borderId = (int)cellStyle.GetCoreXf().GetBorderId();
        ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
        Assert.IsFalse(ctBorder.IsSetLeft());
	}

	public void TestGetSetBorderTop() {
        //default values
        Assert.AreEqual(CellStyle.BORDER_NONE, cellStyle.GetBorderTop());

        int num = stylesTable.GetBorders().Count;
        cellStyle.SetBorderTop(CellStyle.BORDER_MEDIUM);
        Assert.AreEqual(CellStyle.BORDER_MEDIUM, cellStyle.GetBorderTop());
        //a new border has been Added
        Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);
        //id of the Created border
        int borderId = (int)cellStyle.GetCoreXf().GetBorderId();
        Assert.IsTrue(borderId > 0);
        //check Changes in the underlying xml bean
        CTBorder ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
        Assert.AreEqual(STBorderStyle.MEDIUM, ctBorder.GetTop().GetStyle());

        num = stylesTable.GetBorders().Count;
        //setting the same border multiple times should not change borderId
        for (int i = 0; i < 3; i++) {
            cellStyle.SetBorderTop(CellStyle.BORDER_MEDIUM);
            Assert.AreEqual(CellStyle.BORDER_MEDIUM, cellStyle.GetBorderTop());
        }
        Assert.AreEqual(borderId, cellStyle.GetCoreXf().GetBorderId());
        Assert.AreEqual(num, stylesTable.GetBorders().Count);
        Assert.AreSame(ctBorder, stylesTable.GetBorderAt(borderId).GetCTBorder());

        //setting border to none Removes the <top> element
        cellStyle.SetBorderTop(CellStyle.BORDER_NONE);
        Assert.AreEqual(num, stylesTable.GetBorders().Count);
        borderId = (int)cellStyle.GetCoreXf().GetBorderId();
        ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
        Assert.IsFalse(ctBorder.IsSetTop());
	}

	public void TestGetSetBottomBorderColor() {
        //defaults
        Assert.AreEqual(IndexedColors.BLACK.GetIndex(), cellStyle.GetBottomBorderColor());
        Assert.IsNull(cellStyle.GetBottomBorderXSSFColor());

        int num = stylesTable.GetBorders().Count;

        XSSFColor clr;

        //setting indexed color
        cellStyle.SetBottomBorderColor(IndexedColors.BLUE_GREY.GetIndex());
        Assert.AreEqual(IndexedColors.BLUE_GREY.GetIndex(), cellStyle.GetBottomBorderColor());
        clr = cellStyle.GetBottomBorderXSSFColor();
        Assert.IsTrue(clr.GetCTColor().IsSetIndexed());
        Assert.AreEqual(IndexedColors.BLUE_GREY.GetIndex(), clr.GetIndexed());
        //a new border was Added to the styles table
        Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);

        //id of the Created border
        int borderId = (int)cellStyle.GetCoreXf().GetBorderId();
        Assert.IsTrue(borderId > 0);
        //check Changes in the underlying xml bean
        CTBorder ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
        Assert.AreEqual(IndexedColors.BLUE_GREY.GetIndex(), ctBorder.GetBottom().GetColor().GetIndexed());

        //setting XSSFColor
        num = stylesTable.GetBorders().Count;
        clr = new XSSFColor(java.awt.Color.CYAN);
        cellStyle.SetBottomBorderColor(clr);
        Assert.AreEqual(clr.GetCTColor().ToString(), cellStyle.GetBottomBorderXSSFColor().GetCTColor().ToString());
        byte[] rgb = cellStyle.GetBottomBorderXSSFColor().GetRgb();
        Assert.AreEqual(java.awt.Color.CYAN, new java.awt.Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
        //another border was Added to the styles table
        Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);

        //passing null unsets the color
        cellStyle.SetBottomBorderColor(null);
        Assert.IsNull(cellStyle.GetBottomBorderXSSFColor());
    }

	public void TestGetSetTopBorderColor() {
        //defaults
        Assert.AreEqual(IndexedColors.BLACK.GetIndex(), cellStyle.GetTopBorderColor());
        Assert.IsNull(cellStyle.GetTopBorderXSSFColor());

        int num = stylesTable.GetBorders().Count;

        XSSFColor clr;

        //setting indexed color
        cellStyle.SetTopBorderColor(IndexedColors.BLUE_GREY.GetIndex());
        Assert.AreEqual(IndexedColors.BLUE_GREY.GetIndex(), cellStyle.GetTopBorderColor());
        clr = cellStyle.GetTopBorderXSSFColor();
        Assert.IsTrue(clr.GetCTColor().IsSetIndexed());
        Assert.AreEqual(IndexedColors.BLUE_GREY.GetIndex(), clr.GetIndexed());
        //a new border was Added to the styles table
        Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);

        //id of the Created border
        int borderId = (int)cellStyle.GetCoreXf().GetBorderId();
        Assert.IsTrue(borderId > 0);
        //check Changes in the underlying xml bean
        CTBorder ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
        Assert.AreEqual(IndexedColors.BLUE_GREY.GetIndex(), ctBorder.GetTop().GetColor().GetIndexed());

        //setting XSSFColor
        num = stylesTable.GetBorders().Count;
        clr = new XSSFColor(java.awt.Color.CYAN);
        cellStyle.SetTopBorderColor(clr);
        Assert.AreEqual(clr.GetCTColor().ToString(), cellStyle.GetTopBorderXSSFColor().GetCTColor().ToString());
        byte[] rgb = cellStyle.GetTopBorderXSSFColor().GetRgb();
        Assert.AreEqual(java.awt.Color.CYAN, new java.awt.Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
        //another border was Added to the styles table
        Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);

        //passing null unsets the color
        cellStyle.SetTopBorderColor(null);
        Assert.IsNull(cellStyle.GetTopBorderXSSFColor());
	}

	public void TestGetSetLeftBorderColor() {
        //defaults
        Assert.AreEqual(IndexedColors.BLACK.GetIndex(), cellStyle.GetLeftBorderColor());
        Assert.IsNull(cellStyle.GetLeftBorderXSSFColor());

        int num = stylesTable.GetBorders().Count;

        XSSFColor clr;

        //setting indexed color
        cellStyle.SetLeftBorderColor(IndexedColors.BLUE_GREY.GetIndex());
        Assert.AreEqual(IndexedColors.BLUE_GREY.GetIndex(), cellStyle.GetLeftBorderColor());
        clr = cellStyle.GetLeftBorderXSSFColor();
        Assert.IsTrue(clr.GetCTColor().IsSetIndexed());
        Assert.AreEqual(IndexedColors.BLUE_GREY.GetIndex(), clr.GetIndexed());
        //a new border was Added to the styles table
        Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);

        //id of the Created border
        int borderId = (int)cellStyle.GetCoreXf().GetBorderId();
        Assert.IsTrue(borderId > 0);
        //check Changes in the underlying xml bean
        CTBorder ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
        Assert.AreEqual(IndexedColors.BLUE_GREY.GetIndex(), ctBorder.GetLeft().GetColor().GetIndexed());

        //setting XSSFColor
        num = stylesTable.GetBorders().Count;
        clr = new XSSFColor(java.awt.Color.CYAN);
        cellStyle.SetLeftBorderColor(clr);
        Assert.AreEqual(clr.GetCTColor().ToString(), cellStyle.GetLeftBorderXSSFColor().GetCTColor().ToString());
        byte[] rgb = cellStyle.GetLeftBorderXSSFColor().GetRgb();
        Assert.AreEqual(java.awt.Color.CYAN, new java.awt.Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
        //another border was Added to the styles table
        Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);

        //passing null unsets the color
        cellStyle.SetLeftBorderColor(null);
        Assert.IsNull(cellStyle.GetLeftBorderXSSFColor());
	}

	public void TestGetSetRightBorderColor() {
        //defaults
        Assert.AreEqual(IndexedColors.BLACK.GetIndex(), cellStyle.GetRightBorderColor());
        Assert.IsNull(cellStyle.GetRightBorderXSSFColor());

        int num = stylesTable.GetBorders().Count;

        XSSFColor clr;

        //setting indexed color
        cellStyle.SetRightBorderColor(IndexedColors.BLUE_GREY.GetIndex());
        Assert.AreEqual(IndexedColors.BLUE_GREY.GetIndex(), cellStyle.GetRightBorderColor());
        clr = cellStyle.GetRightBorderXSSFColor();
        Assert.IsTrue(clr.GetCTColor().IsSetIndexed());
        Assert.AreEqual(IndexedColors.BLUE_GREY.GetIndex(), clr.GetIndexed());
        //a new border was Added to the styles table
        Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);

        //id of the Created border
        int borderId = (int)cellStyle.GetCoreXf().GetBorderId();
        Assert.IsTrue(borderId > 0);
        //check Changes in the underlying xml bean
        CTBorder ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
        Assert.AreEqual(IndexedColors.BLUE_GREY.GetIndex(), ctBorder.GetRight().GetColor().GetIndexed());

        //setting XSSFColor
        num = stylesTable.GetBorders().Count;
        clr = new XSSFColor(java.awt.Color.CYAN);
        cellStyle.SetRightBorderColor(clr);
        Assert.AreEqual(clr.GetCTColor().ToString(), cellStyle.GetRightBorderXSSFColor().GetCTColor().ToString());
        byte[] rgb = cellStyle.GetRightBorderXSSFColor().GetRgb();
        Assert.AreEqual(java.awt.Color.CYAN, new java.awt.Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
        //another border was Added to the styles table
        Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);

        //passing null unsets the color
        cellStyle.SetRightBorderColor(null);
        Assert.IsNull(cellStyle.GetRightBorderXSSFColor());
	}

	public void TestGetSetFillBackgroundColor() {

        Assert.AreEqual(IndexedColors.AUTOMATIC.GetIndex(), cellStyle.GetFillBackgroundColor());
        Assert.IsNull(cellStyle.GetFillBackgroundXSSFColor());

        XSSFColor clr;

        int num = stylesTable.GetFills().Count;

        //setting indexed color
        cellStyle.SetFillBackgroundColor(IndexedColors.RED.GetIndex());
        Assert.AreEqual(IndexedColors.RED.GetIndex(), cellStyle.GetFillBackgroundColor());
        clr = cellStyle.GetFillBackgroundXSSFColor();
        Assert.IsTrue(clr.GetCTColor().IsSetIndexed());
        Assert.AreEqual(IndexedColors.RED.GetIndex(), clr.GetIndexed());
        //a new fill was Added to the styles table
        Assert.AreEqual(num + 1, stylesTable.GetFills().Count);

        //id of the Created border
        int FillId = (int)cellStyle.GetCoreXf().GetFillId();
        Assert.IsTrue(FillId > 0);
        //check Changes in the underlying xml bean
        CTFill ctFill = stylesTable.GetFillAt(FillId).GetCTFill();
        Assert.AreEqual(IndexedColors.RED.GetIndex(), ctFill.GetPatternFill().GetBgColor().GetIndexed());

        //setting XSSFColor
        num = stylesTable.GetFills().Count;
        clr = new XSSFColor(java.awt.Color.CYAN);
        cellStyle.SetFillBackgroundColor(clr);
        Assert.AreEqual(clr.GetCTColor().ToString(), cellStyle.GetFillBackgroundXSSFColor().GetCTColor().ToString());
        byte[] rgb = cellStyle.GetFillBackgroundXSSFColor().GetRgb();
        Assert.AreEqual(java.awt.Color.CYAN, new java.awt.Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
        //another border was Added to the styles table
        Assert.AreEqual(num + 1, stylesTable.GetFills().Count);

        //passing null unsets the color
        cellStyle.SetFillBackgroundColor(null);
        Assert.IsNull(cellStyle.GetFillBackgroundXSSFColor());
        Assert.AreEqual(IndexedColors.AUTOMATIC.GetIndex(), cellStyle.GetFillBackgroundColor());
	}

	public void TestDefaultStyles() {

		XSSFWorkbook wb1 = new XSSFWorkbook();

		XSSFCellStyle style1 = wb1.CreateCellStyle();
        Assert.AreEqual(IndexedColors.AUTOMATIC.GetIndex(), style1.GetFillBackgroundColor());
        Assert.IsNull(style1.GetFillBackgroundXSSFColor());

        //compatibility with HSSF
        HSSFWorkbook wb2 = new HSSFWorkbook();
        HSSFCellStyle style2 = wb2.CreateCellStyle();
        Assert.AreEqual(style2.GetFillBackgroundColor(), style1.GetFillBackgroundColor());
        Assert.AreEqual(style2.GetFillForegroundColor(), style1.GetFillForegroundColor());
        Assert.AreEqual(style2.GetFillPattern(), style1.GetFillPattern());

        Assert.AreEqual(style2.GetLeftBorderColor(), style1.GetLeftBorderColor());
        Assert.AreEqual(style2.GetTopBorderColor(), style1.GetTopBorderColor());
        Assert.AreEqual(style2.GetRightBorderColor(), style1.GetRightBorderColor());
        Assert.AreEqual(style2.GetBottomBorderColor(), style1.GetBottomBorderColor());

        Assert.AreEqual(style2.GetBorderBottom(), style1.GetBorderBottom());
        Assert.AreEqual(style2.GetBorderLeft(), style1.GetBorderLeft());
        Assert.AreEqual(style2.GetBorderRight(), style1.GetBorderRight());
        Assert.AreEqual(style2.GetBorderTop(), style1.GetBorderTop());
	}


	public void TestGetFillForegroundColor() {

        XSSFWorkbook wb = new XSSFWorkbook();
        StylesTable styles = wb.GetStylesSource();
        Assert.AreEqual(1, wb.GetNumCellStyles());
        Assert.AreEqual(2, styles.GetFills().Count);

        XSSFCellStyle defaultStyle = wb.GetCellStyleAt((short)0);
        Assert.AreEqual(IndexedColors.AUTOMATIC.GetIndex(), defaultStyle.GetFillForegroundColor());
        Assert.AreEqual(null, defaultStyle.GetFillForegroundXSSFColor());
        Assert.AreEqual(CellStyle.NO_FILL, defaultStyle.GetFillPattern());

        XSSFCellStyle customStyle = wb.CreateCellStyle();

        customStyle.SetFillPattern(CellStyle.SOLID_FOREGROUND);
        Assert.AreEqual(CellStyle.SOLID_FOREGROUND, customStyle.GetFillPattern());
        Assert.AreEqual(3, styles.GetFills().Count);

        customStyle.SetFillForegroundColor(IndexedColors.BRIGHT_GREEN.GetIndex());
        Assert.AreEqual(IndexedColors.BRIGHT_GREEN.GetIndex(), customStyle.GetFillForegroundColor());
        Assert.AreEqual(4, styles.GetFills().Count);

        for (int i = 0; i < 3; i++) {
            XSSFCellStyle style = wb.CreateCellStyle();

            style.SetFillPattern(CellStyle.SOLID_FOREGROUND);
            Assert.AreEqual(CellStyle.SOLID_FOREGROUND, style.GetFillPattern());
            Assert.AreEqual(4, styles.GetFills().Count);

            style.SetFillForegroundColor(IndexedColors.BRIGHT_GREEN.GetIndex());
            Assert.AreEqual(IndexedColors.BRIGHT_GREEN.GetIndex(), style.GetFillForegroundColor());
            Assert.AreEqual(4, styles.GetFills().Count);
        }
	}

	public void TestGetFillPattern() {

        Assert.AreEqual(CellStyle.NO_FILL, cellStyle.GetFillPattern());

        int num = stylesTable.GetFills().Count;
        cellStyle.SetFillPattern(CellStyle.SOLID_FOREGROUND);
        Assert.AreEqual(CellStyle.SOLID_FOREGROUND, cellStyle.GetFillPattern());
        Assert.AreEqual(num + 1, stylesTable.GetFills().Count);
        int FillId = (int)cellStyle.GetCoreXf().GetFillId();
        Assert.IsTrue(FillId > 0);
        //check Changes in the underlying xml bean
        CTFill ctFill = stylesTable.GetFillAt(FillId).GetCTFill();
        Assert.AreEqual(STPatternType.SOLID, ctFill.GetPatternFill().GetPatternType());

        //setting the same fill multiple time does not update the styles table
        for (int i = 0; i < 3; i++) {
            cellStyle.SetFillPattern(CellStyle.SOLID_FOREGROUND);
        }
        Assert.AreEqual(num + 1, stylesTable.GetFills().Count);

        cellStyle.SetFillPattern(CellStyle.NO_FILL);
        Assert.AreEqual(CellStyle.NO_FILL, cellStyle.GetFillPattern());
        FillId = (int)cellStyle.GetCoreXf().GetFillId();
        ctFill = stylesTable.GetFillAt(FillId).GetCTFill();
        Assert.IsFalse(ctFill.GetPatternFill().IsSetPatternType());

	}

	public void TestGetFont() {
		Assert.IsNotNull(cellStyle.GetFont());
	}

	public void TestGetSetHidden() {
		Assert.IsFalse(cellStyle.GetHidden());
		cellStyle.SetHidden(true);
		Assert.IsTrue(cellStyle.GetHidden());
		cellStyle.SetHidden(false);
		Assert.IsFalse(cellStyle.GetHidden());
	}

	public void TestGetSetLocked() {
		Assert.IsTrue(cellStyle.GetLocked());
		cellStyle.SetLocked(true);
		Assert.IsTrue(cellStyle.GetLocked());
		cellStyle.SetLocked(false);
		Assert.IsFalse(cellStyle.GetLocked());
	}

	public void TestGetSetIndent() {
		Assert.AreEqual((short)0, cellStyle.GetIndention());
		cellStyle.SetIndention((short)3);
		Assert.AreEqual((short)3, cellStyle.GetIndention());
		cellStyle.SetIndention((short) 13);
		Assert.AreEqual((short)13, cellStyle.GetIndention());
	}

	public void TestGetSetAlignement() {
		Assert.IsNull(cellStyle.GetCellAlignment().GetCTCellAlignment().GetHorizontal());
		Assert.AreEqual(HorizontalAlignment.GENERAL, cellStyle.GetAlignmentEnum());

		cellStyle.SetAlignment(XSSFCellStyle.ALIGN_LEFT);
		Assert.AreEqual(XSSFCellStyle.ALIGN_LEFT, cellStyle.GetAlignment());
		Assert.AreEqual(HorizontalAlignment.LEFT, cellStyle.GetAlignmentEnum());
		Assert.AreEqual(STHorizontalAlignment.LEFT, cellStyle.GetCellAlignment().GetCTCellAlignment().GetHorizontal());

		cellStyle.SetAlignment(HorizontalAlignment.JUSTIFY);
		Assert.AreEqual(XSSFCellStyle.ALIGN_JUSTIFY, cellStyle.GetAlignment());
		Assert.AreEqual(HorizontalAlignment.JUSTIFY, cellStyle.GetAlignmentEnum());
		Assert.AreEqual(STHorizontalAlignment.JUSTIFY, cellStyle.GetCellAlignment().GetCTCellAlignment().GetHorizontal());

		cellStyle.SetAlignment(HorizontalAlignment.CENTER);
		Assert.AreEqual(XSSFCellStyle.ALIGN_CENTER, cellStyle.GetAlignment());
		Assert.AreEqual(HorizontalAlignment.CENTER, cellStyle.GetAlignmentEnum());
		Assert.AreEqual(STHorizontalAlignment.CENTER, cellStyle.GetCellAlignment().GetCTCellAlignment().GetHorizontal());
	}

	public void TestGetSetVerticalAlignment() {
		Assert.AreEqual(VerticalAlignment.BOTTOM, cellStyle.GetVerticalAlignmentEnum());
		Assert.AreEqual(XSSFCellStyle.VERTICAL_BOTTOM, cellStyle.GetVerticalAlignment());
		Assert.IsNull(cellStyle.GetCellAlignment().GetCTCellAlignment().GetVertical());

		cellStyle.SetVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		Assert.AreEqual(XSSFCellStyle.VERTICAL_CENTER, cellStyle.GetVerticalAlignment());
		Assert.AreEqual(VerticalAlignment.CENTER, cellStyle.GetVerticalAlignmentEnum());
		Assert.AreEqual(STVerticalAlignment.CENTER, cellStyle.GetCellAlignment().GetCTCellAlignment().GetVertical());

		cellStyle.SetVerticalAlignment(XSSFCellStyle.VERTICAL_JUSTIFY);
		Assert.AreEqual(XSSFCellStyle.VERTICAL_JUSTIFY, cellStyle.GetVerticalAlignment());
		Assert.AreEqual(VerticalAlignment.JUSTIFY, cellStyle.GetVerticalAlignmentEnum());
		Assert.AreEqual(STVerticalAlignment.JUSTIFY, cellStyle.GetCellAlignment().GetCTCellAlignment().GetVertical());
	}

	public void TestGetSetWrapText() {
		Assert.IsFalse(cellStyle.GetWrapText());
		cellStyle.SetWrapText(true);
		Assert.IsTrue(cellStyle.GetWrapText());
		cellStyle.SetWrapText(false);
        Assert.IsFalse(cellStyle.GetWrapText());
	}

	/**
	 * Cloning one XSSFCellStyle onto Another, same XSSFWorkbook
	 */
	public void TestCloneStyleSameWB() {
      XSSFWorkbook wb = new XSSFWorkbook();
      Assert.AreEqual(1, wb.GetNumberOfFonts());
      
      XSSFFont fnt = wb.CreateFont();
      fnt.SetFontName("TestingFont");
      Assert.AreEqual(2, wb.GetNumberOfFonts());
      
      XSSFCellStyle orig = wb.CreateCellStyle();
      orig.SetAlignment(HSSFCellStyle.ALIGN_RIGHT);
      orig.SetFont(fnt);
      orig.SetDataFormat((short)18);
      
      Assert.IsTrue(HSSFCellStyle.ALIGN_RIGHT == orig.GetAlignment());
      Assert.IsTrue(fnt == orig.GetFont());
      Assert.IsTrue(18 == orig.GetDataFormat());
      
      XSSFCellStyle clone = wb.CreateCellStyle();
      Assert.IsFalse(HSSFCellStyle.ALIGN_RIGHT == Clone.GetAlignment());
      Assert.IsFalse(fnt == Clone.GetFont());
      Assert.IsFalse(18 == Clone.GetDataFormat());
      
      Clone.CloneStyleFrom(orig);
      Assert.IsTrue(HSSFCellStyle.ALIGN_RIGHT == Clone.GetAlignment());
      Assert.IsTrue(fnt == Clone.GetFont());
      Assert.IsTrue(18 == Clone.GetDataFormat());
      Assert.AreEqual(2, wb.GetNumberOfFonts());
	}
	/**
	 * Cloning one XSSFCellStyle onto Another, different XSSFWorkbooks
	 */
	public void TestCloneStyleDiffWB() {
       XSSFWorkbook wbOrig = new XSSFWorkbook();
       Assert.AreEqual(1, wbOrig.GetNumberOfFonts());
       Assert.AreEqual(0, wbOrig.GetStylesSource().GetNumberFormats().Count);
       
       XSSFFont fnt = wbOrig.CreateFont();
       fnt.SetFontName("TestingFont");
       Assert.AreEqual(2, wbOrig.GetNumberOfFonts());
       Assert.AreEqual(0, wbOrig.GetStylesSource().GetNumberFormats().Count);
       
       XSSFDataFormat fmt = wbOrig.CreateDataFormat();
       fmt.GetFormat("MadeUpOne");
       fmt.GetFormat("MadeUpTwo");
       
       XSSFCellStyle orig = wbOrig.CreateCellStyle();
       orig.SetAlignment(HSSFCellStyle.ALIGN_RIGHT);
       orig.SetFont(fnt);
       orig.SetDataFormat(fmt.GetFormat("Test##"));
       
       Assert.IsTrue(XSSFCellStyle.ALIGN_RIGHT == orig.GetAlignment());
       Assert.IsTrue(fnt == orig.GetFont());
       Assert.IsTrue(fmt.GetFormat("Test##") == orig.GetDataFormat());
       
       Assert.AreEqual(2, wbOrig.GetNumberOfFonts());
       Assert.AreEqual(3, wbOrig.GetStylesSource().GetNumberFormats().Count);
       
       
       // Now a style on another workbook
       XSSFWorkbook wbClone = new XSSFWorkbook();
       Assert.AreEqual(1, wbClone.GetNumberOfFonts());
       Assert.AreEqual(0, wbClone.GetStylesSource().GetNumberFormats().Count);
       Assert.AreEqual(1, wbClone.GetNumCellStyles());
       
       XSSFDataFormat fmtClone = wbClone.CreateDataFormat();
       XSSFCellStyle clone = wbClone.CreateCellStyle();
       
       Assert.AreEqual(1, wbClone.GetNumberOfFonts());
       Assert.AreEqual(0, wbClone.GetStylesSource().GetNumberFormats().Count);
       
       Assert.IsFalse(HSSFCellStyle.ALIGN_RIGHT == Clone.GetAlignment());
       Assert.IsFalse("TestingFont" == Clone.GetFont().GetFontName());
       
       Clone.CloneStyleFrom(orig);
       
       Assert.AreEqual(2, wbClone.GetNumberOfFonts());
       Assert.AreEqual(2, wbClone.GetNumCellStyles());
       Assert.AreEqual(1, wbClone.GetStylesSource().GetNumberFormats().Count);
       
       Assert.AreEqual(HSSFCellStyle.ALIGN_RIGHT, Clone.GetAlignment());
       Assert.AreEqual("TestingFont", Clone.GetFont().GetFontName());
       Assert.AreEqual(fmtClone.GetFormat("Test##"), Clone.GetDataFormat());
       Assert.IsFalse(fmtClone.GetFormat("Test##") == fmt.GetFormat("Test##"));
       
       // Save it and re-check
       XSSFWorkbook wbReload = XSSFTestDataSamples.WriteOutAndReadBack(wbClone);
       Assert.AreEqual(2, wbReload.GetNumberOfFonts());
       Assert.AreEqual(2, wbReload.GetNumCellStyles());
       Assert.AreEqual(1, wbReload.GetStylesSource().GetNumberFormats().Count);
       
       XSSFCellStyle reload = wbReload.GetCellStyleAt((short)1);
       Assert.AreEqual(HSSFCellStyle.ALIGN_RIGHT, reload.GetAlignment());
       Assert.AreEqual("TestingFont", reload.GetFont().GetFontName());
       Assert.AreEqual(fmtClone.GetFormat("Test##"), reload.GetDataFormat());
       Assert.IsFalse(fmtClone.GetFormat("Test##") == fmt.GetFormat("Test##"));
   }
}




