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

namespace NPOI.HSLF.Model;

using NPOI.ddf.*;
using NPOI.util.LittleEndian;





/**
 * Represents a table in a PowerPoint presentation
 *
 * @author Yegor Kozlov
 */
public class Table : ShapeGroup {

    protected static int BORDER_TOP = 1;
    protected static int BORDER_RIGHT = 2;
    protected static int BORDER_BOTTOM = 3;
    protected static int BORDER_LEFT = 4;

    protected static int BORDERS_ALL = 5;
    protected static int BORDERS_OUTSIDE = 6;
    protected static int BORDERS_INSIDE = 7;
    protected static int BORDERS_NONE = 8;


    protected TableCell[][] cells;

    /**
     * Create a new Table of the given number of rows and columns
     *
     * @param numrows the number of rows
     * @param numcols the number of columns
     */
    public Table(int numrows, int numcols) {
        base();

        if(numrows < 1) throw new ArgumentException("The number of rows must be greater than 1");
        if(numcols < 1) throw new ArgumentException("The number of columns must be greater than 1");

        int x=0, y=0, tblWidth=0, tblHeight=0;
        cells = new TableCell[numrows][numcols];
        for (int i = 0; i < cells.Length; i++) {
            x = 0;
            for (int j = 0; j < cells[i].Length; j++) {
                cells[i][j] = new TableCell(this);
                Rectangle anchor = new Rectangle(x, y, TableCell.DEFAULT_WIDTH, TableCell.DEFAULT_HEIGHT);
                cells[i][j].SetAnchor(anchor);
                x += TableCell.DEFAULT_WIDTH;
            }
            y += TableCell.DEFAULT_HEIGHT;
        }
        tblWidth = x;
        tblHeight = y;
        SetAnchor(new Rectangle(0, 0, tblWidth, tblHeight));

        EscherContainerRecord spCont = (EscherContainerRecord) GetSpContainer().GetChild(0);
        EscherOptRecord opt = new EscherOptRecord();
        opt.SetRecordId((short)0xF122);
        opt.AddEscherProperty(new EscherSimpleProperty((short)0x39F, 1));
        EscherArrayProperty p = new EscherArrayProperty((short)0x43A0, false, null);
        p.SetSizeOfElements(0x0004);
        p.SetNumberOfElementsInArray(numrows);
        p.SetNumberOfElementsInMemory(numrows);
        opt.AddEscherProperty(p);
        List<EscherRecord> lst = spCont.GetChildRecords();
        lst.Add(lst.Count-1, opt);
        spCont.SetChildRecords(lst);
    }

    /**
     * Create a Table object and Initilize it from the supplied Record Container.
     *
     * @param escherRecord <code>EscherSpContainer</code> Container which holds information about this shape
     * @param parent       the parent of the shape
     */
    public Table(EscherContainerRecord escherRecord, Shape parent) {
        base(escherRecord, parent);
    }

    /**
     * Gets a cell
     *
     * @param row the row index (0-based)
     * @param col the column index (0-based)
     * @return the cell
     */
    public TableCell GetCell(int row, int col) {
        return cells[row][col];
    }

    public int GetNumberOfColumns() {
        return cells[0].Length;
    }
    public int GetNumberOfRows() {
        return cells.Length;
    }

    protected void afterInsert(Sheet sh){
        super.afterInsert(sh);

        EscherContainerRecord spCont = (EscherContainerRecord) GetSpContainer().GetChild(0);
        List<EscherRecord> lst = spCont.GetChildRecords();
        EscherOptRecord opt = (EscherOptRecord)lst.Get(lst.Count-2);
        EscherArrayProperty p = (EscherArrayProperty)opt.GetEscherProperty(1);
        for (int i = 0; i < cells.Length; i++) {
            TableCell cell = cells[i][0];
            int rowHeight = cell.GetAnchor().height*MASTER_DPI/POINT_DPI;
            byte[] val = new byte[4];
            LittleEndian.PutInt(val, rowHeight);
            p.SetElement(i, val);
            for (int j = 0; j < cells[i].Length; j++) {
                TableCell c = cells[i][j];
                AddShape(c);

                Line bt = c.GetBorderTop();
                if(bt != null) AddShape(bt);

                Line br = c.GetBorderRight();
                if(br != null) AddShape(br);

                Line bb = c.GetBorderBottom();
                if(bb != null) AddShape(bb);

                Line bl = c.GetBorderLeft();
                if(bl != null) AddShape(bl);

            }
        }

    }

    protected void InitTable(){
        Shape[] sh = GetShapes();
        Arrays.sort(sh, new Comparator(){
            public int Compare( Object o1, Object o2 ) {
                Rectangle anchor1 = ((Shape)o1).GetAnchor();
                Rectangle anchor2 = ((Shape)o2).GetAnchor();
                int delta = anchor1.y - anchor2.y;
                if(delta == 0) delta = anchor1.x - anchor2.x;
                return delta;
            }
        });
        int y0 = -1;
        int maxrowlen = 0;
        ArrayList lst = new ArrayList();
        ArrayList row = null;
        for (int i = 0; i < sh.Length; i++) {
            if(sh[i] is TextShape){
                Rectangle anchor = sh[i].GetAnchor();
                if(anchor.y != y0){
                    y0 = anchor.y;
                    row = new ArrayList();
                    lst.Add(row);
                }
                row.Add(sh[i]);
                maxrowlen = Math.max(maxrowlen, row.Count);
            }
        }
        cells = new TableCell[lst.Count][maxrowlen];
        for (int i = 0; i < lst.Count; i++) {
            row = (ArrayList)lst.Get(i);
            for (int j = 0; j < row.Count; j++) {
                TextShape tx = (TextShape)row.Get(j);
                cells[i][j] = new TableCell(tx.GetSpContainer(), GetParent());
                cells[i][j].SetSheet(tx.Sheet);
            }
        }
    }

    /**
     * Assign the <code>SlideShow</code> this shape belongs to
     *
     * @param sheet owner of this shape
     */
    public void SetSheet(Sheet sheet){
        super.SetSheet(sheet);
        if(cells == null) InitTable();
    }

    /**
     * Sets the row height.
     *
     * @param row the row index (0-based)
     * @param height the height to Set (in pixels)
     */
    public void SetRowHeight(int row, int height){
        int currentHeight = cells[row][0].GetAnchor().height;
        int dy = height - currentHeight;

        for (int i = row; i < cells.Length; i++) {
            for (int j = 0; j < cells[i].Length; j++) {
                Rectangle anchor = cells[i][j].GetAnchor();
                if(i == row) anchor.height = height;
                else anchor.y += dy;
                cells[i][j].SetAnchor(anchor);
            }
        }
        Rectangle tblanchor = GetAnchor();
        tblanchor.height += dy;
        SetAnchor(tblanchor);

    }

    /**
     * Sets the column width.
     *
     * @param col the column index (0-based)
     * @param width the width to Set (in pixels)
     */
    public void SetColumnWidth(int col, int width){
        int currentWidth = cells[0][col].GetAnchor().width;
        int dx = width - currentWidth;
        for (int i = 0; i < cells.Length; i++) {
            Rectangle anchor = cells[i][col].GetAnchor();
            anchor.width = width;
            cells[i][col].SetAnchor(anchor);

            if(col < cells[i].Length - 1) for (int j = col+1; j < cells[i].Length; j++) {
                anchor = cells[i][j].GetAnchor();
                anchor.x += dx;
                cells[i][j].SetAnchor(anchor);
            }
        }
        Rectangle tblanchor = GetAnchor();
        tblanchor.width += dx;
        SetAnchor(tblanchor);
    }

    /**
     * Format the table and apply the specified Line to all cell boundaries,
     * both outside and inside
     *
     * @param line the border line
     */
    public void SetAllBorders(Line line){
        for (int i = 0; i < cells.Length; i++) {
            for (int j = 0; j < cells[i].Length; j++) {
                TableCell cell = cells[i][j];
                cell.SetBorderTop(CloneBorder(line));
                cell.SetBorderLeft(CloneBorder(line));
                if(j == cells[i].Length - 1) cell.SetBorderRight(CloneBorder(line));
                if(i == cells.Length - 1) cell.SetBorderBottom(CloneBorder(line));
            }
        }
    }

    /**
     * Format the outside border using the specified Line object
     *
     * @param line the border line
     */
    public void SetOutsideBorders(Line line){
        for (int i = 0; i < cells.Length; i++) {
            for (int j = 0; j < cells[i].Length; j++) {
                TableCell cell = cells[i][j];

                if(j == 0) cell.SetBorderLeft(CloneBorder(line));
                if(j == cells[i].Length - 1) cell.SetBorderRight(CloneBorder(line));
                else {
                    cell.SetBorderLeft(null);
                    cell.SetBorderLeft(null);
                }

                if(i == 0) cell.SetBorderTop(CloneBorder(line));
                else if(i == cells.Length - 1) cell.SetBorderBottom(CloneBorder(line));
                else {
                    cell.SetBorderTop(null);
                    cell.SetBorderBottom(null);
                }
            }
        }
    }

    /**
     * Format the inside border using the specified Line object
     *
     * @param line the border line
     */
    public void SetInsideBorders(Line line){
        for (int i = 0; i < cells.Length; i++) {
            for (int j = 0; j < cells[i].Length; j++) {
                TableCell cell = cells[i][j];

                if(j != cells[i].Length - 1)
                    cell.SetBorderRight(CloneBorder(line));
                else {
                    cell.SetBorderLeft(null);
                    cell.SetBorderLeft(null);
                }
                if(i != cells.Length - 1) cell.SetBorderBottom(CloneBorder(line));
                else {
                    cell.SetBorderTop(null);
                    cell.SetBorderBottom(null);
                }
            }
        }
    }

    private Line CloneBorder(Line line){
        Line border = CreateBorder();
        border.SetLineWidth(line.GetLineWidth());
        border.SetLineStyle(line.GetLineStyle());
        border.SetLineDashing(line.GetLineDashing());
        border.SetLineColor(line.GetLineColor());
        return border;
    }

    /**
     * Create a border to format this table
     *
     * @return the Created border
     */
    public Line CreateBorder(){
        Line line = new Line(this);

        EscherOptRecord opt = (EscherOptRecord)getEscherChild(line.GetSpContainer(), EscherOptRecord.RECORD_ID);
        SetEscherProperty(opt, EscherProperties.GEOMETRY__SHAPEPATH, -1);
        SetEscherProperty(opt, EscherProperties.GEOMETRY__FILLOK, -1);
        SetEscherProperty(opt, EscherProperties.SHADOWSTYLE__SHADOWOBSURED, 0x20000);
        SetEscherProperty(opt, EscherProperties.THREED__LIGHTFACE, 0x80000);

        return line;
    }
}





