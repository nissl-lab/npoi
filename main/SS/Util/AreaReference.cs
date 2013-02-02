/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.SS.Util
{

    using System;
    using System.Text;
    using System.Collections;

    public class AreaReference
    {

        /** The Char (!) that Separates sheet names from cell references */
        private const char SHEET_NAME_DELIMITER = '!';
        /** The Char (:) that Separates the two cell references in a multi-cell area reference */
        private const char CELL_DELIMITER = ':';
        /** The Char (') used to quote sheet names when they contain special Chars */
        private const char SPECIAL_NAME_DELIMITER = '\'';

        private CellReference _firstCell;
        private CellReference _lastCell;
        private bool _isSingleCell;

        /**
         * Create an area ref from a string representation.  Sheet names containing special Chars should be
         * delimited and escaped as per normal syntax rules for formulas.<br/> 
         * The area reference must be contiguous (i.e. represent a single rectangle, not a Union of rectangles)
         */
        public AreaReference(String reference)
        {
            if (!IsContiguous(reference))
            {
                throw new ArgumentException(
                        "References passed to the AreaReference must be contiguous, " +
                        "use generateContiguous(ref) if you have non-contiguous references");
            }

            String[] parts = SeparateAreaRefs(reference);

            String part0 = parts[0];
            if (parts.Length == 1)
            {
                // TODO - probably shouldn't initialize area ref when text is really a cell ref
                // Need to fix some named range stuff to get rid of this
                _firstCell = new CellReference(part0);

                _lastCell = _firstCell;
                _isSingleCell = true;
                return;
            }
            if (parts.Length != 2)
            {
                throw new ArgumentException("Bad area ref '" + reference + "'");
            }
            String part1 = parts[1];
            if (IsPlainColumn(part0))
            {
                if (!IsPlainColumn(part1))
                {
                    throw new Exception("Bad area ref '" + reference + "'");
                }
                // Special handling for whole-column references
                // Represented internally as x$1 to x$65536
                //  which is the maximum range of rows

                bool firstIsAbs = CellReference.IsPartAbsolute(part0);
                bool lastIsAbs = CellReference.IsPartAbsolute(part1);

                int col0 = CellReference.ConvertColStringToIndex(part0);
                int col1 = CellReference.ConvertColStringToIndex(part1);

                _firstCell = new CellReference(0, col0, true, firstIsAbs);
                _lastCell = new CellReference(0xFFFF, col1, true, lastIsAbs);
                _isSingleCell = false;
                // TODO - whole row refs
            }
            else
            {
                _firstCell = new CellReference(part0);
                _lastCell = new CellReference(part1);
                _isSingleCell = part0.Equals(part1);
            }
        }

        private bool IsPlainColumn(String refPart)
        {
            for (int i = refPart.Length - 1; i >= 0; i--)
            {
                int ch = refPart[i];
                if (ch == '$' && i == 0)
                {
                    continue;
                }
                if (ch < 'A' || ch > 'Z')
                {
                    return false;
                }
            }
            return true;
        }
        public static AreaReference GetWholeRow(String start, String end)
        {
            return new AreaReference("$A" + start + ":$IV" + end);
        }

        public static AreaReference GetWholeColumn(String start, String end)
        {
            return new AreaReference(start + "$1:" + end + "$65536");
        }


        /**
         * Creates an area ref from a pair of Cell References.
         */
        public AreaReference(CellReference topLeft, CellReference botRight)
        {
            //_firstCell = topLeft;
            //_lastCell = botRight;
            //_isSingleCell = false;

            bool swapRows = topLeft.Row > botRight.Row;
            bool swapCols = topLeft.Col > botRight.Col;
            if (swapRows || swapCols)
            {
                int firstRow;
                int lastRow;
                int firstColumn;
                int lastColumn;
                bool firstRowAbs;
                bool lastRowAbs;
                bool firstColAbs;
                bool lastColAbs;
                if (swapRows)
                {
                    firstRow = botRight.Row;
                    firstRowAbs = botRight.IsRowAbsolute;
                    lastRow = topLeft.Row;
                    lastRowAbs = topLeft.IsRowAbsolute;
                }
                else
                {
                    firstRow = topLeft.Row;
                    firstRowAbs = topLeft.IsRowAbsolute;
                    lastRow = botRight.Row;
                    lastRowAbs = botRight.IsRowAbsolute;
                }
                if (swapCols)
                {
                    firstColumn = botRight.Col;
                    firstColAbs = botRight.IsColAbsolute;
                    lastColumn = topLeft.Col;
                    lastColAbs = topLeft.IsColAbsolute;
                }
                else
                {
                    firstColumn = topLeft.Col;
                    firstColAbs = topLeft.IsColAbsolute;
                    lastColumn = botRight.Col;
                    lastColAbs = botRight.IsColAbsolute;
                }
                _firstCell = new CellReference(firstRow, firstColumn, firstRowAbs, firstColAbs);
                _lastCell = new CellReference(lastRow, lastColumn, lastRowAbs, lastColAbs);
            }
            else
            {
                _firstCell = topLeft;
                _lastCell = botRight;
            }
            _isSingleCell = false;
        }

        /**
         * is the reference for a contiguous (i.e.
         *  Unbroken) area, or is it made up of
         *  several different parts?
         * (If it Is, you will need to call
         *  ....
         */
        public static bool IsContiguous(String reference)
        {
            // If there's a sheet name, strip it off
            int sheetRefEnd = reference.IndexOf('!');
            if (sheetRefEnd != -1)
            {
                reference = reference.Substring(sheetRefEnd);
            }

            // Check for the , as a sign of non-coniguous
            if (reference.IndexOf(',') == -1)
            {
                return true;
            }
            return false;
        }

        /**
         * is the reference for a whole-column reference,
         *  such as C:C or D:G ?
         */
        public static bool IsWholeColumnReference(CellReference topLeft, CellReference botRight)
        {
            // These are represented as something like
            //   C$1:C$65535 or D$1:F$0
            // i.e. absolute from 1st row to 0th one
            if (topLeft.Row == 0 && topLeft.IsRowAbsolute &&
                (botRight.Row == -1 || botRight.Row == 65535) && botRight.IsRowAbsolute)
            {
                return true;
            }
            return false;
        }
        public bool IsWholeColumnReference()
        {
            return IsWholeColumnReference(_firstCell, _lastCell);
        }

        /**
         * Takes a non-contiguous area reference, and
         *  returns an array of contiguous area references.
         */
        public static AreaReference[] GenerateContiguous(String reference)
        {
            ArrayList refs = new ArrayList();
            String st = reference;
            string[] token = st.Split(',');
            foreach (string t in token)
            {
                refs.Add(
                        new AreaReference(t)
                );
            }
            return (AreaReference[])refs.ToArray(typeof(AreaReference));
        }

        /**
         * @return <c>false</c> if this area reference involves more than one cell
         */
        public bool IsSingleCell
        {
            get { return _isSingleCell; }
        }

        /**
         * @return the first cell reference which defines this area. Usually this cell is in the upper
         * left corner of the area (but this is not a requirement).
         */
        public CellReference FirstCell
        {
            get { return _firstCell; }
        }

        /**
         * Note - if this area reference refers to a single cell, the return value of this method will
         * be identical to that of <c>GetFirstCell()</c>
         * @return the second cell reference which defines this area.  For multi-cell areas, this is 
         * cell diagonally opposite the 'first cell'.  Usually this cell is in the lower right corner 
         * of the area (but this is not a requirement).
         */
        public CellReference LastCell
        {
            get{return _lastCell;}
        }
        /**
         * Returns a reference to every cell covered by this area
         */
        public CellReference[] GetAllReferencedCells()
        {
            // Special case for single cell reference
            if (_isSingleCell)
            {
                return new CellReference[] { _firstCell, };
            }

            // Interpolate between the two
            int minRow = Math.Min(_firstCell.Row, _lastCell.Row);
            int maxRow = Math.Max(_firstCell.Row, _lastCell.Row);
            int minCol = Math.Min(_firstCell.Col, _lastCell.Col);
            int maxCol = Math.Max(_firstCell.Col, _lastCell.Col);
            String sheetName = _firstCell.SheetName;

            ArrayList refs = new ArrayList();
            for (int row = minRow; row <= maxRow; row++)
            {
                for (int col = minCol; col <= maxCol; col++)
                {
                    CellReference ref1 = new CellReference(sheetName, row, col, _firstCell.IsRowAbsolute, _firstCell.IsColAbsolute);
                    refs.Add(ref1);
                }
            }
            return (CellReference[])refs.ToArray(typeof(CellReference));
        }

        /**
         *  Example return values:
         *    <table border="0" cellpAdding="1" cellspacing="0" summary="Example return values">
         *      <tr><th align='left'>Result</th><th align='left'>Comment</th></tr>
         *      <tr><td>A1:A1</td><td>Single cell area reference without sheet</td></tr>
         *      <tr><td>A1:$C$1</td><td>Multi-cell area reference without sheet</td></tr>
         *      <tr><td>Sheet1!A$1:B4</td><td>Standard sheet name</td></tr>
         *      <tr><td>'O''Brien''s Sales'!B5:C6' </td><td>Sheet name with special Chars</td></tr>
         *    </table>
         * @return the text representation of this area reference as it would appear in a formula.
         */
        public String FormatAsString()
        {
                // Special handling for whole-column references
                if (IsWholeColumnReference())
                {
                    return
                        CellReference.ConvertNumToColString(_firstCell.Col)
                        + ":" +
                        CellReference.ConvertNumToColString(_lastCell.Col);
                }

                StringBuilder sb = new StringBuilder(32);
                sb.Append(_firstCell.FormatAsString());
                if (!_isSingleCell)
                {
                    sb.Append(CELL_DELIMITER);
                    if (_lastCell.SheetName == null)
                    {
                        sb.Append(_lastCell.FormatAsString());
                    }
                    else
                    {
                        // don't want to include the sheet name twice
                        _lastCell.AppendCellReference(sb);
                    }
                }
                return sb.ToString();
        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(this.GetType().Name).Append(" [");
            sb.Append(FormatAsString());
            sb.Append("]");
            return sb.ToString();
        }

        /**
         * Separates Area refs in two parts and returns them as Separate elements in a String array,
         * each qualified with the sheet name (if present)
         * 
         * @return array with one or two elements. never <c>null</c>
         */
        private static String[] SeparateAreaRefs(String reference)
        {
            // TODO - refactor cell reference parsing logic to one place.
            // Current known incarnations: 
            //   FormulaParser.Name
            //   CellReference.SeparateRefParts() 
            //   AreaReference.SeparateAreaRefs() (here)
            //   SheetNameFormatter.format() (inverse)


            int len = reference.Length;
            int delimiterPos = -1;
            bool insideDelimitedName = false;
            for (int i = 0; i < len; i++)
            {
                switch (reference[i])
                {
                    case CELL_DELIMITER:
                        if (!insideDelimitedName)
                        {
                            if (delimiterPos >= 0)
                            {
                                throw new ArgumentException("More than one cell delimiter '"
                                        + CELL_DELIMITER + "' appears in area reference '" + reference + "'");
                            }
                            delimiterPos = i;
                        }
                        continue;
                    case SPECIAL_NAME_DELIMITER:
                    // fall through
                        break;
                    default:
                        continue;
                }
                if (!insideDelimitedName)
                {
                    insideDelimitedName = true;
                    continue;
                }

                if (i >= len - 1)
                {
                    // reference ends with the delimited name. 
                    // Assume names like: "Sheet1!'A1'" are never legal.
                    throw new ArgumentException("Area reference '" + reference
                            + "' ends with special name delimiter '" + SPECIAL_NAME_DELIMITER + "'");
                }
                if (reference[i + 1] == SPECIAL_NAME_DELIMITER)
                {
                    // two consecutive quotes is the escape sequence for a single one
                    i++; // skip this and keep parsing the special name
                }
                else
                {
                    // this is the end of the delimited name
                    insideDelimitedName = false;
                }
            }
            if (delimiterPos < 0)
            {
                return new String[] { reference, };
            }

            String partA = reference.Substring(0, delimiterPos);
            String partB = reference.Substring(delimiterPos + 1);
            if (partB.IndexOf(SHEET_NAME_DELIMITER) >= 0)
            {
                // TODO - are references like "Sheet1!A1:Sheet1:B2" ever valid?  
                // FormulaParser has code to handle that.

                throw new Exception("Unexpected " + SHEET_NAME_DELIMITER
                        + " in second cell reference of '" + reference + "'");
            }

            int plingPos = partA.LastIndexOf(SHEET_NAME_DELIMITER);
            if (plingPos < 0)
            {
                return new String[] { partA, partB, };
            }

            String sheetName = partA.Substring(0, plingPos + 1); // +1 to include delimiter

            return new String[] { partA, sheetName + partB, };
        }
    }

}