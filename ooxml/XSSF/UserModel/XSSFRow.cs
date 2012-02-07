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

namespace NPOI.XSSF.usermodel;

using java.util.Iterator;
using java.util.TreeMap;

using NPOI.SS.SpreadsheetVersion;
using NPOI.SS.usermodel.Cell;
using NPOI.SS.usermodel.CellStyle;
using NPOI.SS.usermodel.Row;
using NPOI.SS.util.CellReference;
using NPOI.UTIL.Internal;
using NPOI.UTIL.POILogFactory;
using NPOI.UTIL.POILogger;
using NPOI.XSSF.model.CalculationChain;
using NPOI.XSSF.model.StylesTable;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CT_Cell;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTRow;

/**
 * High level representation of a row of a spreadsheet.
 */
public class XSSFRow : Row, Comparable<XSSFRow> {
    private static POILogger _logger = POILogFactory.GetLogger(XSSFRow.class);

    /**
     * the xml bean Containing all cell defInitions for this row
     */
    private CTRow _row;

    /**
     * Cells of this row keyed by their column indexes.
     * The TreeMap ensures that the cells are ordered by columnIndex in the ascending order.
     */
    private TreeDictionary<int, XSSFCell> _cells;

    /**
     * the parent sheet
     */
    private XSSFSheet _sheet;

    /**
     * Construct a XSSFRow.
     *
     * @param row the xml bean Containing all cell defInitions for this row.
     * @param sheet the parent sheet.
     */
    @SuppressWarnings("deprecation") //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    protected XSSFRow(CTRow row, XSSFSheet sheet) {
        _row = row;
        _sheet = sheet;
        _cells = new TreeDictionary<int, XSSFCell>();
        for (CT_Cell c : row.GetCArray()) {
            XSSFCell cell = new XSSFCell(this, c);
            _cells.Put(cell.GetColumnIndex(), cell);
            sheet.onReadCell(cell);
        }
    }

    /**
     * Returns the XSSFSheet this row belongs to
     *
     * @return the XSSFSheet that owns this row
     */
    public XSSFSheet GetSheet() {
        return this._sheet;
    }

    /**
     * Cell iterator over the physically defined cells:
     * <blockquote><pre>
     * for (Iterator<Cell> it = row.cellIterator(); it.HasNext(); ) {
     *     Cell cell = it.next();
     *     ...
     * }
     * </pre></blockquote>
     *
     * @return an iterator over cells in this row.
     */
    public Iterator<Cell> cellIterator() {
        return (Iterator<Cell>)(Iterator<? : Cell>)_cells.values().iterator();
    }

    /**
     * Alias for {@link #cellIterator()} to allow  foreach loops:
     * <blockquote><pre>
     * for(Cell cell : row){
     *     ...
     * }
     * </pre></blockquote>
     *
     * @return an iterator over cells in this row.
     */
    public Iterator<Cell> iterator() {
    	return cellIterator();
    }

    /**
     * Compares two <code>XSSFRow</code> objects.  Two rows are equal if they belong to the same worksheet and
     * their row indexes are Equal.
     *
     * @param   row   the <code>XSSFRow</code> to be Compared.
     * @return	the value <code>0</code> if the row number of this <code>XSSFRow</code> is
     * 		equal to the row number of the argument <code>XSSFRow</code>; a value less than
     * 		<code>0</code> if the row number of this this <code>XSSFRow</code> is numerically less
     * 		than the row number of the argument <code>XSSFRow</code>; and a value greater
     * 		than <code>0</code> if the row number of this this <code>XSSFRow</code> is numerically
     * 		 greater than the row number of the argument <code>XSSFRow</code>.
     * @throws ArgumentException if the argument row belongs to a different worksheet
     */
    public int CompareTo(XSSFRow row) {
        int thisVal = this.GetRowNum();
        if(row.GetSheet() != GetSheet()) throw new ArgumentException("The Compared rows must belong to the same XSSFSheet");

        int anotherVal = row.GetRowNum();
        return (thisVal < anotherVal ? -1 : (thisVal == anotherVal ? 0 : 1));
    }

    /**
     * Use this to create new cells within the row and return it.
     * <p>
     * The cell that is returned is a {@link Cell#CELL_TYPE_BLANK}. The type can be Changed
     * either through calling <code>SetCellValue</code> or <code>SetCellType</code>.
     * </p>
     * @param columnIndex - the column number this cell represents
     * @return Cell a high level representation of the Created cell.
     * @throws ArgumentException if columnIndex < 0 or greater than 16384,
     *   the maximum number of columns supported by the SpreadsheetML format (.xlsx)
     */
    public XSSFCell CreateCell(int columnIndex) {
    	return CreateCell(columnIndex, Cell.CELL_TYPE_BLANK);
    }

    /**
     * Use this to create new cells within the row and return it.
     *
     * @param columnIndex - the column number this cell represents
     * @param type - the cell's data type
     * @return XSSFCell a high level representation of the Created cell.
     * @throws ArgumentException if the specified cell type is invalid, columnIndex < 0
     *   or greater than 16384, the maximum number of columns supported by the SpreadsheetML format (.xlsx)
     * @see Cell#CELL_TYPE_BLANK
     * @see Cell#CELL_TYPE_BOOLEAN
     * @see Cell#CELL_TYPE_ERROR
     * @see Cell#CELL_TYPE_FORMULA
     * @see Cell#CELL_TYPE_NUMERIC
     * @see Cell#CELL_TYPE_STRING
     */
    public XSSFCell CreateCell(int columnIndex, int type) {
        CT_Cell CT_Cell;
        XSSFCell prev = _cells.Get(columnIndex);
        if(prev != null){
            CT_Cell = prev.GetCT_Cell();
            CT_Cell.Set(CT_Cell.Factory.newInstance());
        } else {
            CT_Cell = _row.AddNewC();
        }
        XSSFCell xcell = new XSSFCell(this, CT_Cell);
        xcell.SetCellNum(columnIndex);
        if (type != Cell.CELL_TYPE_BLANK) {
        	xcell.SetCellType(type);
        }
        _cells.Put(columnIndex, xcell);
        return xcell;
    }

    /**
     * Returns the cell at the given (0 based) index,
     *  with the {@link NPOI.SS.usermodel.Row.MissingCellPolicy} from the parent Workbook.
     *
     * @return the cell at the given (0 based) index
     */
    public XSSFCell GetCell(int cellnum) {
    	return GetCell(cellnum, _sheet.GetWorkbook().getMissingCellPolicy());
    }

    /**
     * Returns the cell at the given (0 based) index, with the specified {@link NPOI.SS.usermodel.Row.MissingCellPolicy}
     *
     * @return the cell at the given (0 based) index
     * @throws ArgumentException if cellnum < 0 or the specified MissingCellPolicy is invalid
     * @see Row#RETURN_NULL_AND_BLANK
     * @see Row#RETURN_BLANK_AS_NULL
     * @see Row#CREATE_NULL_AS_BLANK
     */
    public XSSFCell GetCell(int cellnum, MissingCellPolicy policy) {
    	if(cellnum < 0) throw new ArgumentException("Cell index must be >= 0");

        XSSFCell cell = (XSSFCell)_cells.Get(cellnum);
    	if(policy == RETURN_NULL_AND_BLANK) {
    		return cell;
    	}
    	if(policy == RETURN_BLANK_AS_NULL) {
    		if(cell == null) return cell;
    		if(cell.GetCellType() == Cell.CELL_TYPE_BLANK) {
    			return null;
    		}
    		return cell;
    	}
    	if(policy == CREATE_NULL_AS_BLANK) {
    		if(cell == null) {
    			return CreateCell((short)cellnum, Cell.CELL_TYPE_BLANK);
    		}
    		return cell;
    	}
    	throw new ArgumentException("Illegal policy " + policy + " (" + policy.id + ")");
    }

    /**
     * Get the number of the first cell Contained in this row.
     *
     * @return short representing the first logical cell in the row,
     *  or -1 if the row does not contain any cells.
     */
    public short GetFirstCellNum() {
    	return (short)(_cells.Count == 0 ? -1 : _cells.firstKey());
    }

    /**
     * Gets the index of the last cell Contained in this row <b>PLUS ONE</b>. The result also
     * happens to be the 1-based column number of the last cell.  This value can be used as a
     * standard upper bound when iterating over cells:
     * <pre>
     * short minColIx = row.GetFirstCellNum();
     * short maxColIx = row.GetLastCellNum();
     * for(short colIx=minColIx; colIx&lt;maxColIx; colIx++) {
     *   XSSFCell cell = row.GetCell(colIx);
     *   if(cell == null) {
     *     continue;
     *   }
     *   //... do something with cell
     * }
     * </pre>
     *
     * @return short representing the last logical cell in the row <b>PLUS ONE</b>,
     *   or -1 if the row does not contain any cells.
     */
    public short GetLastCellNum() {
    	return (short)(_cells.Count == 0 ? -1 : (_cells.lastKey() + 1));
    }

    /**
     * Get the row's height measured in twips (1/20th of a point). If the height is not Set, the default worksheet value is returned,
     * See {@link NPOI.XSSF.usermodel.XSSFSheet#GetDefaultRowHeightInPoints()}
     *
     * @return row height measured in twips (1/20th of a point)
     */
    public short GetHeight() {
        return (short)(GetHeightInPoints()*20);
    }

    /**
     * Returns row height measured in point size. If the height is not Set, the default worksheet value is returned,
     * See {@link NPOI.XSSF.usermodel.XSSFSheet#GetDefaultRowHeightInPoints()}
     *
     * @return row height measured in point size
     * @see NPOI.XSSF.usermodel.XSSFSheet#GetDefaultRowHeightInPoints()
     */
    public float GetHeightInPoints() {
        if (this._row.IsSetHt()) {
            return (float) this._row.GetHt();
        }
        return _sheet.GetDefaultRowHeightInPoints();
    }

    /**
     *  Set the height in "twips" or  1/20th of a point.
     *
     * @param height the height in "twips" or  1/20th of a point. <code>-1</code>  reSets to the default height
     */
    public void SetHeight(short height) {
        if (height == -1) {
            if (_row.IsSetHt()) _row.unSetHt();
            if (_row.IsSetCustomHeight()) _row.unSetCustomHeight();
        } else {
            _row.SetHt((double) height / 20);
            _row.SetCustomHeight(true);

        }
    }

    /**
     * Set the row's height in points.
     *
     * @param height the height in points. <code>-1</code>  reSets to the default height
     */
    public void SetHeightInPoints(float height) {
	    SetHeight((short)(height == -1 ? -1 : (height*20)));
    }

    /**
     * Gets the number of defined cells (NOT number of cells in the actual row!).
     * That is to say if only columns 0,4,5 have values then there would be 3.
     *
     * @return int representing the number of defined cells in the row.
     */
    public int GetPhysicalNumberOfCells() {
    	return _cells.Count;
    }

    /**
     * Get row number this row represents
     *
     * @return the row number (0 based)
     */
    public int GetRowNum() {
        return (int) (_row.GetR() - 1);
    }

    /**
     * Set the row number of this row.
     *
     * @param rowIndex  the row number (0-based)
     * @throws ArgumentException if rowNum < 0 or greater than 1048575
     */
    public void SetRowNum(int rowIndex) {
        int maxrow = SpreadsheetVersion.EXCEL2007.GetLastRowIndex();
        if (rowIndex < 0 || rowIndex > maxrow) {
            throw new ArgumentException("Invalid row number (" + rowIndex
                    + ") outside allowable range (0.." + maxrow + ")");
        }
        _row.SetR(rowIndex + 1);
    }

    /**
     * Get whether or not to display this row with 0 height
     *
     * @return - height is zero or not.
     */
    public bool GetZeroHeight() {
    	return this._row.GetHidden();
    }

    /**
     * Set whether or not to display this row with 0 height
     *
     * @param height  height is zero or not.
     */
    public void SetZeroHeight(bool height) {
    	this._row.SetHidden(height);

    }

    /**
     * Is this row formatted? Most aren't, but some rows
     *  do have whole-row styles. For those that do, you
     *  can get the formatting from {@link #GetRowStyle()}
     */
    public bool IsFormatted() {
        return _row.IsSetS();
    }
    /**
     * Returns the whole-row cell style. Most rows won't
     *  have one of these, so will return null. Call
     *  {@link #isFormatted()} to check first.
     */
    public XSSFCellStyle GetRowStyle() {
       if(!isFormatted()) return null;
       
       StylesTable stylesSource = GetSheet().getWorkbook().getStylesSource();
       if(stylesSource.GetNumCellStyles() > 0) {
           return stylesSource.GetStyleAt((int)_row.getS());
       } else {
          return null;
       }
    }
    
    /**
     * Applies a whole-row cell styling to the row.
     * If the value is null then the style information is Removed,
     *  causing the cell to used the default workbook style.
     */
    public void SetRowStyle(CellStyle style) {
        if(style == null) {
           if(_row.IsSetS()) {
              _row.unSetS();
              _row.unSetCustomFormat();
           }
        } else {
            StylesTable styleSource = GetSheet().getWorkbook().getStylesSource();
            
            XSSFCellStyle xStyle = (XSSFCellStyle)style;
            xStyle.verifyBelongsToStylesSource(styleSource);

            long idx = styleSource.PutStyle(xStyle);
            _row.SetS(idx);
            _row.SetCustomFormat(true);
        }
    }
    
    /**
     * Remove the Cell from this row.
     *
     * @param cell the cell to remove
     */
    public void RemoveCell(Cell cell) {
        if (cell.GetRow() != this) {
            throw new ArgumentException("Specified cell does not belong to this row");
        }

        XSSFCell xcell = (XSSFCell)cell;
        if(xcell.IsPartOfArrayFormulaGroup()) {
            xcell.NotifyArrayFormulaChanging();
        }
        if(cell.GetCellType() == Cell.CELL_TYPE_FORMULA) {
           _sheet.GetWorkbook().onDeleteFormula(xcell);
        }
        _cells.Remove(cell.getColumnIndex());
    }

    /**
     * Returns the underlying CTRow xml bean Containing all cell defInitions in this row
     *
     * @return the underlying CTRow xml bean
     */
    
    public CTRow GetCTRow(){
    	return _row;
    }

    /**
     * Fired when the document is written to an output stream.
     *
     * @see NPOI.XSSF.usermodel.XSSFSheet#Write(java.io.OutputStream) ()
     */
    protected void onDocumentWrite(){
        // check if cells in the CTRow are ordered
        bool IsOrdered = true;
        if(_row.sizeOfCArray() != _cells.Count) isOrdered = false;
        else {
            int i = 0;
            for (XSSFCell cell : _cells.values()) {
                CT_Cell c1 = cell.GetCT_Cell();
                CT_Cell c2 = _row.GetCArray(i++); 

                String r1 = c1.GetR();
                String r2 = c2.GetR();
                if (!(r1==null ? r2==null : r1.Equals(r2))){
                    isOrdered = false;
                    break;
                }
            }
        }

        if(!isOrdered){
            CT_Cell[] cArray = new CT_Cell[_cells.Count];
            int i = 0;
            for (XSSFCell c : _cells.values()) {
                cArray[i++] = c.GetCT_Cell();
            }
            _row.SetCArray(cArray);
        }
    }

    /**
     * @return formatted xml representation of this row
     */
    
    public String ToString(){
        return _row.ToString();
    }

    /**
     * update cell references when Shifting rows
     *
     * @param n the number of rows to move
     */
    protected void Shift(int n) {
        int rownum = GetRowNum() + n;
        CalculationChain calcChain = _sheet.GetWorkbook().getCalculationChain();
        int sheetId = (int)_sheet.sheet.GetSheetId();
        String msg = "Row[rownum="+GetRowNum()+"] Contains cell(s) included in a multi-cell array formula. " +
                "You cannot change part of an array.";
        for(Cell c : this){
            XSSFCell cell = (XSSFCell)c;
            if(cell.IsPartOfArrayFormulaGroup()){
                cell.NotifyArrayFormulaChanging(msg);
            }

            //remove the reference in the calculation chain
            if(calcChain != null) calcChain.RemoveItem(sheetId, cell.GetReference());

            CT_Cell CT_Cell = cell.GetCT_Cell();
            String r = new CellReference(rownum, cell.GetColumnIndex()).formatAsString();
            CT_Cell.SetR(r);
        }
        SetRowNum(rownum);
    }
}

