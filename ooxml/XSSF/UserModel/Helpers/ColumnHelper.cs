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

namespace NPOI.XSSF.usermodel.helpers;

using java.util.Arrays;
using java.util.Iterator;
using java.text.AttributedString;
using java.text.NumberFormat;
using java.text.DecimalFormat;
using java.awt.font.TextLayout;
using java.awt.font.FontRenderContext;
using java.awt.font.TextAttribute;
using java.awt.geom.AffineTransform;

using NPOI.SS.usermodel.CellStyle;
using NPOI.SS.util.CellRangeAddress;
using NPOI.XSSF.util.CTColComparator;
using NPOI.XSSF.util.NumericRanges;
using NPOI.XSSF.usermodel.*;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

/**
 * Helper class for dealing with the Column Settings on
 *  a CTWorksheet (the data part of a sheet).
 * Note - within POI, we use 0 based column indexes, but
 *  the column defInitions in the XML are 1 based!
 */
public class ColumnHelper {

    private CTWorksheet worksheet;
    private CTCols newCols;

    public ColumnHelper(CTWorksheet worksheet) {
        base();
        this.worksheet = worksheet;
        cleanColumns();
    }

    @SuppressWarnings("deprecation") //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    public void cleanColumns() {
        this.newCols = CTCols.Factory.newInstance();
        CTCols[] colsArray = worksheet.GetColsArray();
        int i = 0;
        for (i = 0; i < colsArray.Length; i++) {
            CTCols cols = colsArray[i];
            CTCol[] colArray = cols.GetColArray();
            for (int y = 0; y < colArray.Length; y++) {
                CTCol col = colArray[y];
                newCols = AddCleanColIntoCols(newCols, col);
            }
        }
        for (int y = i - 1; y >= 0; y--) {
            worksheet.RemoveCols(y);
        }
        worksheet.AddNewCols();
        worksheet.SetColsArray(0, newCols);
    }

    @SuppressWarnings("deprecation") //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    public static void sortColumns(CTCols newCols) {
        CTCol[] colArray = newCols.GetColArray();
        Arrays.sort(colArray, new CTColComparator());
        newCols.SetColArray(colArray);
    }

    public CTCol CloneCol(CTCols cols, CTCol col) {
        CTCol newCol = cols.AddNewCol();
        newCol.SetMin(col.GetMin());
        newCol.SetMax(col.GetMax());
        SetColumnAttributes(col, newCol);
        return newCol;
    }

    /**
     * Returns the Column at the given 0 based index
     */
    public CTCol GetColumn(long index, bool splitColumns) {
    	return GetColumn1Based(index+1, splitColumns);
    }
    /**
     * Returns the Column at the given 1 based index.
     * POI default is 0 based, but the file stores
     *  as 1 based.
     */
    public CTCol GetColumn1Based(long index1, bool splitColumns) {
        CTCols colsArray = worksheet.GetColsArray(0);
		for (int i = 0; i < colsArray.sizeOfColArray(); i++) {
            CTCol colArray = colsArray.GetColArray(i);
			if (colArray.GetMin() <= index1 && colArray.GetMax() >= index1) {
				if (splitColumns) {
					if (colArray.GetMin() < index1) {
						insertCol(colsArray, colArray.GetMin(), (index1 - 1), new CTCol[]{colArray});
					}
					if (colArray.GetMax() > index1) {
						insertCol(colsArray, (index1 + 1), colArray.GetMax(), new CTCol[]{colArray});
					}
					colArray.SetMin(index1);
					colArray.SetMax(index1);
				}
                return colArray;
            }
        }
        return null;
    }

    public CTCols AddCleanColIntoCols(CTCols cols, CTCol col) {
        bool colOverlaps = false;
        for (int i = 0; i < cols.sizeOfColArray(); i++) {
            CTCol ithCol = cols.GetColArray(i);
            long[] range1 = { ithCol.GetMin(), ithCol.GetMax() };
            long[] range2 = { col.GetMin(), col.GetMax() };
            long[] overlappingRange = NumericRanges.GetOverlappingRange(range1,
                    range2);
            int overlappingType = NumericRanges.GetOverlappingType(range1,
                    range2);
            // different behavior required for each of the 4 different
            // overlapping types
            if (overlappingType == NumericRanges.OVERLAPS_1_MINOR) {
                ithCol.SetMax(overlappingRange[0] - 1);
                CTCol rangeCol = insertCol(cols, overlappingRange[0],
                        overlappingRange[1], new CTCol[] { ithCol, col });
                i++;
                CTCol newCol = insertCol(cols, (overlappingRange[1] + 1), col
                        .GetMax(), new CTCol[] { col });
                i++;
            } else if (overlappingType == NumericRanges.OVERLAPS_2_MINOR) {
                ithCol.SetMin(overlappingRange[1] + 1);
                CTCol rangeCol = insertCol(cols, overlappingRange[0],
                        overlappingRange[1], new CTCol[] { ithCol, col });
                i++;
                CTCol newCol = insertCol(cols, col.GetMin(),
                        (overlappingRange[0] - 1), new CTCol[] { col });
                i++;
            } else if (overlappingType == NumericRanges.OVERLAPS_2_WRAPS) {
                SetColumnAttributes(col, ithCol);
                if (col.GetMin() != ithCol.GetMin()) {
                    CTCol newColBefore = insertCol(cols, col.GetMin(), (ithCol
                            .GetMin() - 1), new CTCol[] { col });
                    i++;
                }
                if (col.GetMax() != ithCol.GetMax()) {
                    CTCol newColAfter = insertCol(cols, (ithCol.GetMax() + 1),
                            col.GetMax(), new CTCol[] { col });
                    i++;
                }
            } else if (overlappingType == NumericRanges.OVERLAPS_1_WRAPS) {
                if (col.GetMin() != ithCol.GetMin()) {
                    CTCol newColBefore = insertCol(cols, ithCol.GetMin(), (col
                            .GetMin() - 1), new CTCol[] { ithCol });
                    i++;
                }
                if (col.GetMax() != ithCol.GetMax()) {
                    CTCol newColAfter = insertCol(cols, (col.GetMax() + 1),
                            ithCol.GetMax(), new CTCol[] { ithCol });
                    i++;
                }
                ithCol.SetMin(overlappingRange[0]);
                ithCol.SetMax(overlappingRange[1]);
                SetColumnAttributes(col, ithCol);
            }
            if (overlappingType != NumericRanges.NO_OVERLAPS) {
                colOverlaps = true;
            }
        }
        if (!colOverlaps) {
            CTCol newCol = CloneCol(cols, col);
        }
        sortColumns(cols);
        return cols;
    }

    /*
     * Insert a new CTCol at position 0 into cols, Setting min=min, max=max and
     * copying all the colsWithAttributes array cols attributes into newCol
     */
    private CTCol insertCol(CTCols cols, long min, long max,            
        CTCol[] colsWithAttributes) {
        if(!columnExists(cols,min,max)){
                CTCol newCol = cols.insertNewCol(0);
                newCol.SetMin(min);
                newCol.SetMax(max);
                for (CTCol col : colsWithAttributes) {
                        SetColumnAttributes(col, newCol);
                }
                return newCol;
        }
        return null;
    }

    /**
     * Does the column at the given 0 based index exist
     *  in the supplied list of column defInitions?
     */
    public bool columnExists(CTCols cols, long index) {
    	return columnExists1Based(cols, index+1);
    }
    private bool columnExists1Based(CTCols cols, long index1) {
        for (int i = 0; i < cols.sizeOfColArray(); i++) {
            if (cols.GetColArray(i).GetMin() == index1) {
                return true;
            }
        }
        return false;
    }

    public void SetColumnAttributes(CTCol fromCol, CTCol toCol) {
    	if(fromCol.isSetBestFit()) toCol.SetBestFit(fromCol.GetBestFit());
        if(fromCol.isSetCustomWidth()) toCol.SetCustomWidth(fromCol.GetCustomWidth());
        if(fromCol.isSetHidden()) toCol.SetHidden(fromCol.GetHidden());
        if(fromCol.isSetStyle()) toCol.SetStyle(fromCol.GetStyle());
        if(fromCol.isSetWidth()) toCol.SetWidth(fromCol.GetWidth());
        if(fromCol.isSetCollapsed()) toCol.SetCollapsed(fromCol.GetCollapsed());
        if(fromCol.isSetPhonetic()) toCol.SetPhonetic(fromCol.GetPhonetic());
        if(fromCol.isSetOutlineLevel()) toCol.SetOutlineLevel(fromCol.GetOutlineLevel());
        toCol.SetCollapsed(fromCol.isSetCollapsed());
    }

    public void SetColBestFit(long index, bool bestFit) {
        CTCol col = GetOrCreateColumn1Based(index+1, false);
        col.SetBestFit(bestFit);
    }
    public void SetCustomWidth(long index, bool bestFit) {
        CTCol col = GetOrCreateColumn1Based(index+1, true);
        col.SetCustomWidth(bestFit);
    }

    public void SetColWidth(long index, double width) {
        CTCol col = GetOrCreateColumn1Based(index+1, true);
        col.SetWidth(width);
    }

    public void SetColHidden(long index, bool hidden) {
        CTCol col = GetOrCreateColumn1Based(index+1, true);
        col.SetHidden(hidden);
    }

    /**
     * Return the CTCol at the given (0 based) column index,
     *  creating it if required.
     */
    protected CTCol GetOrCreateColumn1Based(long index1, bool splitColumns) {
        CTCol col = GetColumn1Based(index1, splitColumns);
        if (col == null) {
            col = worksheet.GetColsArray(0).AddNewCol();
            col.SetMin(index1);
            col.SetMax(index1);
        }
        return col;
    }

	public void SetColDefaultStyle(long index, CellStyle style) {
		SetColDefaultStyle(index, style.GetIndex());
	}
	
	public void SetColDefaultStyle(long index, int styleId) {
		CTCol col = GetOrCreateColumn1Based(index+1, true);
		col.SetStyle(styleId);
	}
	
	// Returns -1 if no column is found for the given index
	public int GetColDefaultStyle(long index) {
		if (GetColumn(index, false) != null) {
			return (int) GetColumn(index, false).GetStyle();
		}
		return -1;
	}

	private bool columnExists(CTCols cols, long min, long max) {
	    for (int i = 0; i < cols.sizeOfColArray(); i++) {
	        if (cols.GetColArray(i).GetMin() == min && cols.GetColArray(i).GetMax() == max) {
	            return true;
	        }
	    }
	    return false;
	}
	
	public int GetIndexOfColumn(CTCols cols, CTCol col) {
	    for (int i = 0; i < cols.sizeOfColArray(); i++) {
	        if (cols.GetColArray(i).GetMin() == col.GetMin() && cols.GetColArray(i).GetMax() == col.GetMax()) {
	            return i;
	        }
	    }
	    return -1;
	}
}


