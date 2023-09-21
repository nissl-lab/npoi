/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.Formula
{

    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Text;
    using System.Text.RegularExpressions;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Constant;
    using NPOI.SS.Formula.Function;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;

    /// <summary>
    /// Specific exception thrown when a supplied formula does not Parse properly.
    ///  Primarily used by test cases when testing for specific parsing exceptions.
    /// </summary>
    [Serializable]
    public class FormulaParseException : Exception
    {
        /// <summary>
        ///This class was given package scope until it would become Clear that it is useful to general client code.
        /// </summary>
        /// <param name="msg"></param>
        public FormulaParseException(String msg)
            : base(msg)
        {

        }
    }
    /*
     * This class Parses a formula string into a List of Tokens in RPN order.
     * Inspired by
     *           Lets Build a Compiler, by Jack Crenshaw
     * BNF for the formula expression is :
     * <expression> ::= <term> [<addop> <term>]*
     * <term> ::= <factor>  [ <mulop> <factor> ]*
     * <factor> ::= <number> | (<expression>) | <cellRef> | <function>
     * <function> ::= <functionName> ([expression [, expression]*])
     */
    public class FormulaParser
    {
        private String _formulaString;
        private int _formulaLength;
        private int _pointer;

        private ParseNode _rootNode;

        private const char TAB = '\t';// HSSF + XSSF
        private const char CR = '\r';  // Normally just XSSF
        private const char LF = '\n';  // Normally just XSSF

        /**
         * Lookahead Character.
         * Gets value '\0' when the input string is exhausted
         */
        private char look;

        /**
         * Tracks whether the run of whitespace preceeding "look" could be an
         * intersection operator.  See GetChar.
         */
        private bool _inIntersection = false;

        private IFormulaParsingWorkbook _book;
        private static SpreadsheetVersion _ssVersion;

        private int _sheetIndex;
        private int _rowIndex; // 0-based

        /**
         * Create the formula Parser, with the string that is To be
         *  Parsed against the supplied workbook.
         * A later call the Parse() method To return ptg list in
         *  rpn order, then call the GetRPNPtg() To retrive the
         *  Parse results.
         * This class is recommended only for single threaded use.
         *
         * If you only have a usermodel.HSSFWorkbook, and not a
         *  model.Workbook, then use the convenience method on
         *  usermodel.HSSFFormulaEvaluator
         */
        public FormulaParser(String formula, IFormulaParsingWorkbook book, int sheetIndex, int rowIndex)
        {
            _formulaString = formula;
            _pointer = 0;
            this._book = book;

            _ssVersion = book == null ? SpreadsheetVersion.EXCEL97 : book.GetSpreadsheetVersion();
            _formulaLength = _formulaString.Length;
            _sheetIndex = sheetIndex;
            _rowIndex = rowIndex;
        }


        /**
         * Parse a formula into a array of tokens
         * Side effect: creates name (Workbook.createName) if formula contains unrecognized names (names are likely UDFs)
         *
         * @param formula	 the formula to parse
         * @param workbook	the parent workbook
         * @param formulaType the type of the formula, see {@link FormulaType}
         * @param sheetIndex  the 0-based index of the sheet this formula belongs to.
         * @param rowIndex  - the related cell's row index in 0-based form (-1 if the formula is not cell related)
	     *                     used to handle structured references that have the "#This Row" quantifier.
         * The sheet index is required to resolve sheet-level names. <code>-1</code> means that
         * the scope of the name will be ignored and  the parser will match names only by name
         *
         * @return array of parsed tokens
         * @throws FormulaParseException if the formula is unparsable
         */
        public static Ptg[] Parse(String formula, IFormulaParsingWorkbook workbook, FormulaType formulaType, int sheetIndex, int rowIndex)
        {
            FormulaParser fp = new FormulaParser(formula, workbook, sheetIndex, rowIndex);
            fp.Parse();
            return fp.GetRPNPtg(formulaType);
        }

        public static Ptg[] Parse(String formula, IFormulaParsingWorkbook workbook, FormulaType formulaType, int sheetIndex)
        {
            return Parse(formula, workbook, formulaType, sheetIndex, -1);
        }

        /**
         * Parse a structured reference. Converts the structured
         *  reference to the area that represent it.
         *
         * @param tableText - The structured reference text
         * @param workbook - the parent workbook
         * @param rowIndex - the 0-based cell's row index ( used to handle "#This Row" quantifiers )
         * @return the area that being represented by the structured reference.
         */
        public static Area3DPxg ParseStructuredReference(String tableText, IFormulaParsingWorkbook workbook, int rowIndex)
        {
            Ptg[] arr = FormulaParser.Parse(tableText, workbook, 0, 0, rowIndex);
            if (arr.Length != 1 || !(arr[0] is Area3DPxg) ) {
                throw new InvalidOperationException("Illegal structured reference");
            }
            return (Area3DPxg)arr[0];
        }

        /** Read New Character From Input Stream */
        private void GetChar()
        {
            // The intersection operator is a space.  We track whether the run of 
            // whitespace preceeding "look" counts as an intersection operator.  
            if (IsWhite(look))
            {
                if (look == ' ')
                {
                    _inIntersection = true;
                }
            }
            else
            {
                _inIntersection = false;
            }
            // Check To see if we've walked off the end of the string.
            if (_pointer > _formulaLength)
            {
                throw new Exception("too far");
            }
            if (_pointer < _formulaLength)
            {
                look = _formulaString[_pointer];
            }
            else
            {
                // Just return if so and reset 'look' To something To keep
                // SkipWhitespace from spinning
                look = (char)0;
                _inIntersection = false;
            }
            _pointer++;
            //Console.WriteLine("Got char: "+ look);
        }

        /** Report What Was Expected */
        private Exception expected(String s)
        {
            String msg;

            if (look == '=' && _formulaString.Substring(0, _pointer - 1).Trim().Length < 1)
            {
                msg = "The specified formula '" + _formulaString
                    + "' starts with an equals sign which is not allowed.";
            }
            else
            {
                msg = "Parse error near char " + (_pointer - 1) + " '" + look + "'"
                    + " in specified formula '" + _formulaString + "'. Expected "
                    + s;
            }
            return new FormulaParseException(msg);
        }

        /** Recognize an Alpha Character */
        private static bool IsAlpha(char c)
        {
            return Char.IsLetter(c) || c == '$' || c == '_';
        }

        /** Recognize a Decimal Digit */
        private static bool IsDigit(char c)
        {
            return Char.IsDigit(c);
        }

        /** Recognize an Alphanumeric */
        private static bool IsAlNum(char c)
        {
            return IsAlpha(c) || IsDigit(c);
        }

        /** Recognize White Space */
        private static bool IsWhite(char c)
        {
            return c == ' ' || c == TAB || c == CR || c == LF;
        }

        /** Skip Over Leading White Space */
        private void SkipWhite()
        {
            while (IsWhite(look))
            {
                GetChar();
            }
        }

        /**
         *  Consumes the next input character if it is equal To the one specified otherwise throws an
         *  unchecked exception. This method does <b>not</b> consume whitespace (before or after the
         *  matched character).
         */
        private void Match(char x)
        {
            if (look != x)
            {
                throw expected("'" + x + "'");
            }
            GetChar();
        }
        private String ParseUnquotedIdentifier()
        {
            if (look == '\'')
            {
                throw expected("unquoted identifier");
            }
            StringBuilder sb = new StringBuilder();
            while (Char.IsLetterOrDigit(look) || look == '.')
            {
                sb.Append(look);
                GetChar();
            }
            if (sb.Length < 1)
            {
                return null;
            }

            return sb.ToString();
        }
        /** Get a Number */
        private String GetNum()
        {
            StringBuilder value = new StringBuilder();

            while (IsDigit(this.look))
            {
                value.Append(this.look);
                GetChar();
            }
            return value.Length == 0 ? null : value.ToString();
        }

        private ParseNode ParseRangeExpression()
        {
            ParseNode result = ParseRangeable();
            bool hasRange = false;
            while (look == ':')
            {
                int pos = _pointer;
                GetChar();
                ParseNode nextPart = ParseRangeable();
                // Note - no range simplification here. An expr like "A1:B2:C3:D4:E5" should be
                // grouped into area ref pairs like: "(A1:B2):(C3:D4):E5"
                // Furthermore, Excel doesn't seem to simplify
                // expressions like "Sheet1!A1:Sheet1:B2" into "Sheet1!A1:B2"

                CheckValidRangeOperand("LHS", pos, result);
                CheckValidRangeOperand("RHS", pos, nextPart);

                ParseNode[] children = { result, nextPart, };
                result = new ParseNode(RangePtg.instance, children);
                hasRange = true;
            }
            if (hasRange)
            {
                return AugmentWithMemPtg(result);
            }
            return result;
        }
        private static ParseNode AugmentWithMemPtg(ParseNode root)
        {
            Ptg memPtg;
            if (NeedsMemFunc(root))
            {
                memPtg = new MemFuncPtg(root.EncodedSize);
            }
            else
            {
                memPtg = new MemAreaPtg(root.EncodedSize);
            }
            return new ParseNode(memPtg, root);
        }
        /**
         * From OOO doc: "Whenever one operand of the reference subexpression is a function,
         *  a defined name, a 3D reference, or an external reference (and no error occurs),
         *  a tMemFunc token is used"
         *
         */
        private static bool NeedsMemFunc(ParseNode root)
        {
            Ptg token = root.GetToken();
            if (token is AbstractFunctionPtg)
            {
                return true;
            }
            if (token is IExternSheetReferenceToken)
            { // 3D refs
                return true;
            }
            if (token is NamePtg || token is NameXPtg)
            { // 3D refs
                return true;
            }

            if (token is OperationPtg || token is ParenthesisPtg)
            {
                // expect RangePtg, but perhaps also UnionPtg, IntersectionPtg etc
                foreach (ParseNode child in root.GetChildren())
                {
                    if (NeedsMemFunc(child))
                    {
                        return true;
                    }
                }
                return false;
            }
            if (token is OperandPtg)
            {
                return false;
            }
            if (token is OperationPtg)
            {
                return true;
            }

            return false;
        }

        /**
         *
         * @return <c>true</c> if the specified character may be used in a defined name
         */
        private static bool IsValidDefinedNameChar(char ch)
        {
            if (Char.IsLetterOrDigit(ch))
            {
                return true;
            }
            switch (ch)
            {
                case '.':
                case '_':
                case '?':
                case '\\': // of all things
                    return true;
            }
            return false;
        }
        /**
         * @param currentParsePosition used to format a potential error message
         */
        private void CheckValidRangeOperand(String sideName, int currentParsePosition, ParseNode pn)
        {
            if (!IsValidRangeOperand(pn))
            {
                throw new FormulaParseException("The " + sideName
                        + " of the range operator ':' at position "
                        + currentParsePosition + " is not a proper reference.");
            }
        }
        /**
          * @return false if sub-expression represented the specified ParseNode definitely
          * cannot appear on either side of the range (':') operator
          */
        private bool IsValidRangeOperand(ParseNode a)
        {
            Ptg tkn = a.GetToken();
            // Note - order is important for these instance-of checks
            if (tkn is OperandPtg)
            {
                // notably cell refs and area refs
                return true;
            }

            // next 2 are special cases of OperationPtg
            if (tkn is AbstractFunctionPtg)
            {
                AbstractFunctionPtg afp = (AbstractFunctionPtg)tkn;
                byte returnClass = afp.DefaultOperandClass;
                return Ptg.CLASS_REF == returnClass;
            }
            if (tkn is ValueOperatorPtg)
            {
                return false;
            }
            if (tkn is OperationPtg)
            {
                return true;
            }

            // one special case of ControlPtg
            if (tkn is ParenthesisPtg)
            {
                // parenthesis Ptg should have only one child
                return IsValidRangeOperand(a.GetChildren()[0]);
            }

            // one special case of ScalarConstantPtg
            if (tkn == ErrPtg.REF_INVALID)
            {
                return true;
            }

            // All other ControlPtgs and ScalarConstantPtgs cannot be used with ':'
            return false;
        }




        /**
         * Parses area refs (things which could be the operand of ':') and simple factors
         * Examples
         * <pre>
         *   A$1
         *   $A$1 :  $B1
         *   A1 .......	C2
         *   Sheet1 !$A1
         *   a..b!A1
         *   'my sheet'!A1
         *   .my.sheet!A1
         *   'my sheet':'my alt sheet'!A1
         *   .my.sheet1:.my.sheet2!$B$2
         *   my.named..range.
         *   'my sheet'!my.named.range
         *   .my.sheet!my.named.range
         *   foo.bar(123.456, "abc")
         *   123.456
         *   "abc"
         *   true
         *   [Foo.xls]!$A$1
         *   [Foo.xls]'my sheet'!$A$1
         *   [Foo.xls]!my.named.range
         * </pre>
         *
         */
        private ParseNode ParseRangeable()
        {
            SkipWhite();
            int savePointer = _pointer;
            SheetIdentifier sheetIden = ParseSheetName();
            if (sheetIden == null)
            {
                ResetPointer(savePointer);
            }
            else
            {
                SkipWhite();
                savePointer = _pointer;
            }

            SimpleRangePart part1 = ParseSimpleRangePart();

            if (part1 == null)
            {
                if (sheetIden != null)
                {
                    if (look == '#')
                    {  // error ref like MySheet!#REF!
                        return new ParseNode(ErrPtg.ValueOf(ParseErrorLiteral()));
                    }
                    else
                    {
                        // Is it a named range?
                        String name = ParseAsName();
                        if (name.Length == 0)
                        {
                            throw new FormulaParseException("Cell reference or Named Range "
                                    + "expected after sheet name at index " + _pointer + ".");
                        }
                        Ptg nameXPtg = _book.GetNameXPtg(name, sheetIden);
                        if (nameXPtg == null)
                        {
                            throw new FormulaParseException("Specified name '" + name +
                                    "' for sheet " + sheetIden.AsFormulaString() + " not found");
                        }
                        return new ParseNode(nameXPtg);
                    }
                }
                return ParseNonRange(savePointer);
            }





            bool whiteAfterPart1 = IsWhite(look);
            if (whiteAfterPart1)
            {
                SkipWhite();
            }

            if (look == ':')
            {
                int colonPos = _pointer;
                GetChar();
                SkipWhite();
                SimpleRangePart part2 = ParseSimpleRangePart();
                if (part2 != null && !part1.IsCompatibleForArea(part2))
                {
                    // second part is not compatible with an area ref e.g. S!A1:S!B2
                    // where S might be a sheet name (that looks like a column name)

                    part2 = null;
                }
                if (part2 == null)
                {
                    // second part is not compatible with an area ref e.g. A1:OFFSET(B2, 1, 2)
                    // reset and let caller use explicit range operator
                    ResetPointer(colonPos);
                    if (!part1.IsCell)
                    {
                        String prefix;
                        if (sheetIden == null)
                        {
                            prefix = "";
                        }
                        else
                        {
                            prefix = "'" + sheetIden.SheetId.Name + '!';
                        }
                        throw new FormulaParseException(prefix + part1.Rep + "' is not a proper reference.");
                    }
                    return CreateAreaRefParseNode(sheetIden, part1, part2);
                }
                return CreateAreaRefParseNode(sheetIden, part1, part2);
            }

            if (look == '.')
            {
                GetChar();
                int dotCount = 1;
                while (look == '.')
                {
                    dotCount++;
                    GetChar();
                }
                bool whiteBeforePart2 = IsWhite(look);

                SkipWhite();
                SimpleRangePart part2 = ParseSimpleRangePart();
                String part1And2 = _formulaString.Substring(savePointer - 1, _pointer - savePointer);
                if (part2 == null)
                {
                    if (sheetIden != null)
                    {
                        throw new FormulaParseException("Complete area reference expected after sheet name at index "
                                + _pointer + ".");
                    }
                    return ParseNonRange(savePointer);
                }


                if (whiteAfterPart1 || whiteBeforePart2)
                {
                    if (part1.IsRowOrColumn || part2.IsRowOrColumn)
                    {
                        // "A .. B" not valid syntax for "A:B"
                        // and there's no other valid expression that fits this grammar
                        throw new FormulaParseException("Dotted range (full row or column) expression '"
                                + part1And2 + "' must not contain whitespace.");
                    }
                    return CreateAreaRefParseNode(sheetIden, part1, part2);
                }

                if (dotCount == 1 && part1.IsRow && part2.IsRow)
                {
                    // actually, this is looking more like a number
                    return ParseNonRange(savePointer);
                }

                if (part1.IsRowOrColumn || part2.IsRowOrColumn)
                {
                    if (dotCount != 2)
                    {
                        throw new FormulaParseException("Dotted range (full row or column) expression '" + part1And2
                                + "' must have exactly 2 dots.");
                    }
                }
                return CreateAreaRefParseNode(sheetIden, part1, part2);
            }
            if (part1.IsCell && IsValidCellReference(part1.Rep))
            {
                return CreateAreaRefParseNode(sheetIden, part1, null);
            }
            if (sheetIden != null)
            {
                throw new FormulaParseException("Second part of cell reference expected after sheet name at index "
                        + _pointer + ".");
            }

            return ParseNonRange(savePointer);
        }

        private static String specHeaders = "Headers";
        private static String specAll = "All";
        private static String specData = "Data";
        private static String specTotals = "Totals";
        private static String specThisRow = "This Row";

        /**
         * Parses a structured reference, returns it as area reference.
         * Examples:
         * <pre>
         * Table1[col]
         * Table1[[#Totals],[col]]
         * Table1[#Totals]
         * Table1[#All]
         * Table1[#Data]
         * Table1[#Headers]
         * Table1[#Totals]
         * Table1[#This Row]
         * Table1[[#All],[col]]
         * Table1[[#Headers],[col]]
         * Table1[[#Totals],[col]]
         * Table1[[#All],[col1]:[col2]]
         * Table1[[#Data],[col1]:[col2]]
         * Table1[[#Headers],[col1]:[col2]]
         * Table1[[#Totals],[col1]:[col2]]
         * Table1[[#Headers],[#Data],[col2]]
         * Table1[[#This Row], [col1]]
         * Table1[ [col1]:[col2] ]
         * </pre>
         * @param tableName
         * @return
         */
        private ParseNode ParseStructuredReference(String tableName)
        {

            if (!(_ssVersion.Equals(SpreadsheetVersion.EXCEL2007)))
            {
                throw new FormulaParseException("Strctured references work only on XSSF (Excel 2007)!");
            }
            ITable tbl = _book.GetTable(tableName);
            if (tbl == null)
            {
                throw new FormulaParseException("Illegal table name: '" + tableName + "'");
            }
            String sheetName = tbl.SheetName;

            int startCol = tbl.StartColIndex;
            int endCol = tbl.EndColIndex;
            int startRow = tbl.StartRowIndex;
            int endRow = tbl.EndRowIndex;

            // Do NOT return before done reading all the structured reference tokens from the input stream.
            // Throwing exceptions is okay.
            int savePtr0 = _pointer;
            GetChar();

            bool isTotalsSpec = false;
            bool isThisRowSpec = false;
            bool isDataSpec = false;
            bool isHeadersSpec = false;
            bool isAllSpec = false;
            int nSpecQuantifiers = 0; // The number of special quantifiers
            int savePtr1;
            while (true)
            {
                savePtr1 = _pointer;
                String specName = ParseAsSpecialQuantifier();
                if (specName == null)
                {
                    ResetPointer(savePtr1);
                    break;
                }
                if (specName.Equals(specAll))
                {
                    isAllSpec = true;
                }
                else if (specName.Equals(specData))
                {
                    isDataSpec = true;
                }
                else if (specName.Equals(specHeaders))
                {
                    isHeadersSpec = true;
                }
                else if (specName.Equals(specThisRow))
                {
                    isThisRowSpec = true;
                }
                else if (specName.Equals(specTotals))
                {
                    isTotalsSpec = true;
                }
                else
                {
                    throw new FormulaParseException("Unknown special quantifier " + specName);
                }
                nSpecQuantifiers++;
                if (look == ',')
                {
                    GetChar();
                }
                else
                {
                    break;
                }
            }
            bool isThisRow = false;
            SkipWhite();
            if (look == '@')
            {
                isThisRow = true;
                GetChar();
            }
            // parse column quantifier
            String startColumnName = null;
            String endColumnName = null;
            int nColQuantifiers = 0;
            savePtr1 = _pointer;
            startColumnName = ParseAsColumnQuantifier();
            if (startColumnName == null)
            {
                ResetPointer(savePtr1);
            }
            else
            {
                nColQuantifiers++;
                if (look == ',')
                {
                    throw new FormulaParseException("The formula " + _formulaString + "is illegal: you should not use ',' with column quantifiers");
                }
                else if (look == ':')
                {
                    GetChar();
                    endColumnName = ParseAsColumnQuantifier();
                    nColQuantifiers++;
                    if (endColumnName == null)
                    {
                        throw new FormulaParseException("The formula " + _formulaString + "is illegal: the string after ':' must be column quantifier");
                    }
                }
            }

            if (nColQuantifiers == 0 && nSpecQuantifiers == 0)
            {
                ResetPointer(savePtr0);
                savePtr0 = _pointer;
                startColumnName = ParseAsColumnQuantifier();
                if (startColumnName != null)
                {
                    nColQuantifiers++;
                }
                else
                {
                    ResetPointer(savePtr0);
                    String name = ParseAsSpecialQuantifier();
                    if (name != null)
                    {
                        if (name.Equals(specAll))
                        {
                            isAllSpec = true;
                        }
                        else if (name.Equals(specData))
                        {
                            isDataSpec = true;
                        }
                        else if (name.Equals(specHeaders))
                        {
                            isHeadersSpec = true;
                        }
                        else if (name.Equals(specThisRow))
                        {
                            isThisRowSpec = true;
                        }
                        else if (name.Equals(specTotals))
                        {
                            isTotalsSpec = true;
                        }
                        else
                        {
                            throw new FormulaParseException("Unknown special quantifier " + name);
                        }
                        nSpecQuantifiers++;
                    }
                    else
                    {
                        throw new FormulaParseException("The formula " + _formulaString + " is illegal");
                    }
                }
            }
            else
            {
                Match(']');
            }

            // Done reading from input stream
            // Ok to return now

            if (isTotalsSpec && !tbl.IsHasTotalsRow)
            {
                return new ParseNode(ErrPtg.REF_INVALID);
            }
            if ((isThisRow || isThisRowSpec) && (_rowIndex < startRow || endRow < _rowIndex))
            {
                // structured reference is trying to reference a row above or below the table with [#This Row] or [@]
                if (_rowIndex >= 0)
                {
                    return new ParseNode(ErrPtg.VALUE_INVALID);
                }
                else
                {
                    throw new FormulaParseException(
                            "Formula contained [#This Row] or [@] structured reference but this row < 0. " +
                            "Row index must be specified for row-referencing structured references.");
                }
            }

            int actualStartRow = startRow;
            int actualEndRow = endRow;
            int actualStartCol = startCol;
            int actualEndCol = endCol;
            if (nSpecQuantifiers > 0)
            {
                //Selecting rows
                if (nSpecQuantifiers == 1 && isAllSpec)
                {
                    //do nothing
                }
                else if (isDataSpec && isHeadersSpec)
                {
                    if (tbl.IsHasTotalsRow)
                    {
                        actualEndRow = endRow - 1;
                    }
                }
                else if (isDataSpec && isTotalsSpec)
                {
                    actualStartRow = startRow + 1;
                }
                else if (nSpecQuantifiers == 1 && isDataSpec)
                {
                    actualStartRow = startRow + 1;
                    if (tbl.IsHasTotalsRow)
                    {
                        actualEndRow = endRow - 1;
                    }
                }
                else if (nSpecQuantifiers == 1 && isHeadersSpec)
                {
                    actualEndRow = actualStartRow;
                }
                else if (nSpecQuantifiers == 1 && isTotalsSpec)
                {
                    actualStartRow = actualEndRow;
                }
                else if ((nSpecQuantifiers == 1 && isThisRowSpec) || isThisRow)
                {
                    actualStartRow = _rowIndex; //The rowNum is 0 based
                    actualEndRow = _rowIndex;
                }
                else
                {
                    throw new FormulaParseException("The formula " + _formulaString + " is illegal");
                }
            }
            else
            {
                if (isThisRow)
                { // there is a @
                    actualStartRow = _rowIndex; //The rowNum is 0 based
                    actualEndRow = _rowIndex;
                }
                else
                { // Really no special quantifiers
                    actualStartRow++;
                }
            }
            //Selecting cols
            if (nColQuantifiers == 2)
            {
                if (startColumnName == null || endColumnName == null)
                {
                    throw new InvalidOperationException("Fatal error");
                }
                int startIdx = tbl.FindColumnIndex(startColumnName);
                int endIdx = tbl.FindColumnIndex(endColumnName);
                if (startIdx == -1 || endIdx == -1)
                {
                    throw new FormulaParseException("One of the columns " + startColumnName + ", " + endColumnName + " doesn't exist in table " + tbl.Name);
                }
                actualStartCol = startCol + startIdx;
                actualEndCol = startCol + endIdx;

            }
            else if (nColQuantifiers == 1 && !isThisRow)
            {
                if (startColumnName == null)
                {
                    throw new InvalidOperationException("Fatal error");
                }
                int idx = tbl.FindColumnIndex(startColumnName);
                if (idx == -1)
                {
                    throw new FormulaParseException("The column " + startColumnName + " doesn't exist in table " + tbl.Name);
                }
                actualStartCol = startCol + idx;
                actualEndCol = actualStartCol;
            }
            CellReference topLeft = new CellReference(actualStartRow, actualStartCol);
            CellReference bottomRight = new CellReference(actualEndRow, actualEndCol);
            SheetIdentifier sheetIden = new SheetIdentifier(null, new NameIdentifier(sheetName, true));
            Ptg ptg = _book.Get3DReferencePtg(new AreaReference(topLeft, bottomRight), sheetIden);
            return new ParseNode(ptg);
        }

        /**
         * Tries to parse the next as column - can contain whitespace
         * Caller should save pointer.
         * @return
        */
        private String ParseAsColumnQuantifier()
        {
            if (look != '[')
            {
                return null;
            }
            GetChar();
            if (look == '#')
            {
                return null;
            }
            if (look == '@')
            {
                GetChar();
            }
            StringBuilder name = new StringBuilder();
            while (look != ']')
            {
                name.Append(look);
                GetChar();
            }
            Match(']');
            return name.ToString();
        }
        /**
         * Tries to parse the next as special quantifier
         * Caller should save pointer.
         * @return
         */
        private String ParseAsSpecialQuantifier()
        {
            if (look != '[')
            {
                return null;
            }
            GetChar();
            if (look != '#')
            {
                return null;
            }
            GetChar();
            String name = ParseAsName();
            if (name.Equals("This"))
            {
                name = name + ' ' + ParseAsName();
            }
            Match(']');
            return name;
        }

        /**
          * Parses simple factors that are not primitive ranges or range components
          * i.e. '!', ':'(and equiv '...') do not appear
          * Examples
          * <pre>
          *   my.named...range.
          *   foo.bar(123.456, "abc")
          *   123.456
          *   "abc"
          *   true
          * </pre>
          */
        private ParseNode ParseNonRange(int savePointer)
        {
            ResetPointer(savePointer);

            if (Char.IsDigit(look))
            {
                return new ParseNode(ParseNumber());
            }
            if (look == '"')
            {
                return new ParseNode(new StringPtg(ParseStringLiteral()));
            }
            // from now on we can only be dealing with non-quoted identifiers
            // which will either be named ranges or functions
            String name = ParseAsName();
            if (look == '(')
            {
                return Function(name);
            }

            //TODO Livshen's code
            if (look == '[')
            {
                return ParseStructuredReference(name);
            }
            //TODO End of Livshen's code

            if (name.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || name.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
            {
                return new ParseNode(new BoolPtg(name.ToUpper()));
            }
            if (_book == null)
            {
                // Only test cases omit the book (expecting it not to be needed)
                throw new InvalidOperationException("Need book to evaluate name '" + name + "'");
            }
            IEvaluationName evalName = _book.GetName(name, _sheetIndex);
            if (evalName == null)
            {
                throw new FormulaParseException("Specified named range '"
                        + name + "' does not exist in the current workbook.");
            }
            if (evalName.IsRange)
            {
                return new ParseNode(evalName.CreatePtg());
            }
            // TODO - what about NameX ?
            throw new FormulaParseException("Specified name '"
                    + name + "' is not a range as expected.");
        }


        private String ParseAsName()
        {
            StringBuilder sb = new StringBuilder();

            // defined names may begin with a letter or underscore  or backslash
            if (!char.IsLetter(look) && look != '_' && look != '\\')
            {
                throw expected("number, string, defined name, or data table");
            }
            while (IsValidDefinedNameChar(look))
            {
                sb.Append(look);
                GetChar();
            }
            SkipWhite();

            return sb.ToString();
        }

        private int GetSheetExtIx(SheetIdentifier sheetIden)
        {
            int extIx;
            if (sheetIden == null)
            {
                extIx = int.MaxValue;
            }
            else
            {
                String sName = sheetIden.SheetId.Name;
                if (sheetIden.BookName == null)
                {
                    extIx = _book.GetExternalSheetIndex(sName);
                }
                else
                {
                    extIx = _book.GetExternalSheetIndex(sheetIden.BookName, sName);
                }
            }
            return extIx;
        }

        /**
         *
         * @param sheetIden may be <code>null</code>
         * @param part1
         * @param part2 may be <code>null</code>
         */
        private ParseNode CreateAreaRefParseNode(SheetIdentifier sheetIden, SimpleRangePart part1,
                SimpleRangePart part2)
        {
            Ptg ptg;
            if (part2 == null)
            {
                CellReference cr = part1.CellReference;
                if (sheetIden == null)
                {
                    ptg = new RefPtg(cr);
                }
                else
                {
                    ptg = _book.Get3DReferencePtg(cr, sheetIden);
                }
            }
            else
            {
                AreaReference areaRef = CreateAreaRef(part1, part2);

                if (sheetIden == null)
                {
                    ptg = new AreaPtg(areaRef);
                }
                else
                {
                    ptg = _book.Get3DReferencePtg(areaRef, sheetIden);
                }
            }
            return new ParseNode(ptg);
        }
        private static AreaReference CreateAreaRef(SimpleRangePart part1, SimpleRangePart part2)
        {
            if (!part1.IsCompatibleForArea(part2))
            {
                throw new FormulaParseException("has incompatible parts: '"
                        + part1.Rep + "' and '" + part2.Rep + "'.");
            }
            if (part1.IsRow)
            {
                return AreaReference.GetWholeRow(_ssVersion, part1.Rep, part2.Rep);
            }
            if (part1.IsColumn)
            {
                return AreaReference.GetWholeColumn(_ssVersion, part1.Rep, part2.Rep);
            }
            return new AreaReference(part1.CellReference, part2.CellReference);
        }


        /**
          * Parses out a potential LHS or RHS of a ':' intended to produce a plain AreaRef.  Normally these are
          * proper cell references but they could also be row or column refs like "$AC" or "10"
          * @return <code>null</code> (and leaves {@link #_pointer} unchanged if a proper range part does not parse out
          */
        private SimpleRangePart ParseSimpleRangePart()
        {
            int ptr = _pointer - 1; // TODO avoid StringIndexOutOfBounds
            bool hasDigits = false;
            bool hasLetters = false;
            while (ptr < _formulaLength)
            {
                char ch = _formulaString[ptr];
                if (Char.IsDigit(ch))
                {
                    hasDigits = true;
                }
                else if (Char.IsLetter(ch))
                {
                    hasLetters = true;
                }
                else if (ch == '$' || ch == '_')    //fix poi bug 49725
                {
                    //do nothing
                }
                else
                {
                    break;
                }
                ptr++;
            }
            if (ptr <= _pointer - 1)
            {
                return null;
            }
            ReadOnlySpan<char> rep = _formulaString.AsSpan(_pointer - 1, ptr - _pointer + 1);

            if (!CellReferenceParser.TryParseCellReference(rep, out _, out var column, out _, out var row))
            {
                return null;
            }
            // Check range bounds against grid max
            if (hasLetters && hasDigits)
            {
                if (!IsValidCellReference(rep))
                {
                    return null;
                }
            }
            else if (hasLetters)
            {
                if (!CellReference.IsColumnWithinRange(column, _ssVersion))
                {
                    return null;
                }
            }
            else if (hasDigits)
            {
                if (!CellReferenceParser.TryParsePositiveInt32Fast(row, out int i))
                {
                    return null;
                }
                if (i < 1 || i > _ssVersion.MaxRows)
                {
                    return null;
                }
            }
            else
            {
                // just dollars ? can this happen?
                return null;
            }


            ResetPointer(ptr + 1); // stepping forward
            return new SimpleRangePart(rep.ToString(), hasLetters, hasDigits);
        }



        /**
         * 
         * "A1", "B3" -> "A1:B3"   
         * "sheet1!A1", "B3" -> "sheet1!A1:B3"
         * 
         * @return <c>null</c> if the range expression cannot / shouldn't be reduced.
         */
        private static Ptg ReduceRangeExpression(Ptg ptgA, Ptg ptgB)
        {
            if (!(ptgB is RefPtg))
            {
                // only when second ref is simple 2-D ref can the range 
                // expression be converted To an area ref
                return null;
            }
            RefPtg refB = (RefPtg)ptgB;

            if (ptgA is RefPtg)
            {
                RefPtg refA = (RefPtg)ptgA;
                return new AreaPtg(refA.Row, refB.Row, refA.Column, refB.Column,
                        refA.IsRowRelative, refB.IsRowRelative, refA.IsColRelative, refB.IsColRelative);
            }
            if (ptgA is Ref3DPtg)
            {
                Ref3DPtg refA = (Ref3DPtg)ptgA;
                return new Area3DPtg(refA.Row, refB.Row, refA.Column, refB.Column,
                        refA.IsRowRelative, refB.IsRowRelative, refA.IsColRelative, refB.IsColRelative,
                        refA.ExternSheetIndex);
            }
            // Note - other operand types (like AreaPtg) which probably can't evaluate 
            // do not cause validation errors at Parse time
            return null;
        }
        /**
         * A1, $A1, A$1, $A$1, A, 1
         */
        private class SimpleRangePart
        {
            public enum PartType
            {
                Cell, Row, Column
            }

            public static PartType Get(bool hasLetters, bool hasDigits)
            {
                if (hasLetters)
                {
                    return hasDigits ? PartType.Cell : PartType.Column;
                }
                if (!hasDigits)
                {
                    throw new ArgumentException("must have either letters or numbers");
                }
                return PartType.Row;
            }

            private PartType _type;
            private String _rep;

            public SimpleRangePart(String rep, bool hasLetters, bool hasNumbers)
            {
                _rep = rep;
                _type = Get(hasLetters, hasNumbers);
            }

            public bool IsCell
            {
                get
                {
                    return _type == PartType.Cell;
                }
            }

            public bool IsRowOrColumn
            {
                get
                {
                    return _type != PartType.Cell;
                }
            }


            public CellReference CellReference
            {
                get
                {
                    if (_type != PartType.Cell)
                    {
                        throw new InvalidOperationException("Not applicable to this type");
                    }
                    return new CellReference(_rep);
                }
            }

            public bool IsColumn
            {
                get
                {
                    return _type == PartType.Column;
                }
            }

            public bool IsRow
            {
                get
                {
                    return _type == PartType.Row;
                }
            }

            public String Rep
            {
                get
                {
                    return _rep;
                }
            }

            /**
             * @return <c>true</c> if the two range parts can be combined in an
             * {@link AreaPtg} ( Note - the explicit range operator (:) may still be valid
             * when this method returns <c>false</c> )
             */
            public bool IsCompatibleForArea(SimpleRangePart part2)
            {
                return _type == part2._type;
            }

            public override String ToString()
            {
                StringBuilder sb = new StringBuilder(64);
                sb.Append(this.GetType().Name).Append(" [");
                sb.Append(_rep);
                sb.Append("]");
                return sb.ToString();
            }
        }
        /**
         * Note - caller should reset {@link #_pointer} upon <code>null</code> result
         * @return The sheet name as an identifier <code>null</code> if '!' is not found in the right place
         */
        private SheetIdentifier ParseSheetName()
        {

            String bookName;
            if (look == '[')
            {
                StringBuilder sb = new StringBuilder();
                GetChar();
                while (look != ']')
                {
                    sb.Append(look);
                    GetChar();
                }
                GetChar();
                bookName = sb.ToString();
            }
            else
            {
                bookName = null;
            }

            if (look == '\'')
            {
                StringBuilder sb = new StringBuilder();

                Match('\'');
                bool done = look == '\'';
                while (!done)
                {
                    sb.Append(look);
                    GetChar();
                    if (look == '\'')
                    {
                        Match('\'');
                        done = look != '\'';
                    }
                }

                NameIdentifier iden = new NameIdentifier(sb.ToString(), true);
                // quoted identifier - can't concatenate anything more
                SkipWhite();
                if (look == '!')
                {
                    GetChar();
                    return new SheetIdentifier(bookName, iden);
                }
                // See if it's a multi-sheet range, eg Sheet1:Sheet3!A1
                if (look == ':')
                {
                    return ParseSheetRange(bookName, iden);
                }
                return null;
            }

            // unquoted sheet names must start with underscore or a letter
            if (look == '_' || Char.IsLetter(look))
            {
                StringBuilder sb = new StringBuilder();
                // can concatenate idens with dots
                while (IsUnquotedSheetNameChar(look))
                {
                    sb.Append(look);
                    GetChar();
                }
                NameIdentifier iden = new NameIdentifier(sb.ToString(), false);
                SkipWhite();
                if (look == '!')
                {
                    GetChar();
                    return new SheetIdentifier(bookName, iden);
                }
                // See if it's a multi-sheet range, eg Sheet1:Sheet3!A1
                if (look == ':')
                {
                    return ParseSheetRange(bookName, iden);
                }
                return null;
            }

            if (look == '!' && bookName != null)
            {
                // Raw book reference, without a sheet
                GetChar();
                return new SheetIdentifier(bookName, null);
            }
            return null;
        }

        /**
         * If we have something that looks like [book]Sheet1: or 
         *  Sheet1, see if it's actually a range eg Sheet1:Sheet2!
         */
        private SheetIdentifier ParseSheetRange(String bookname, NameIdentifier sheet1Name)
        {
            GetChar();
            SheetIdentifier sheet2 = ParseSheetName();
            if (sheet2 != null)
            {
                return new SheetRangeIdentifier(bookname, sheet1Name, sheet2._sheetIdentifier);
            }
            return null;
        }
        /**
          * very similar to {@link SheetNameFormatter#isSpecialChar(char)}
          */
        private bool IsUnquotedSheetNameChar(char ch)
        {
            if (Char.IsLetterOrDigit(ch))
            {
                return true;
            }
            switch (ch)
            {
                case '.': // dot is OK
                case '_': // underscore is OK
                    return true;
            }
            return false;
        }
        private void ResetPointer(int ptr)
        {
            _pointer = ptr;
            if (_pointer <= _formulaLength)
            {
                look = _formulaString[_pointer - 1];
            }
            else
            {
                // Just return if so and reset 'look' to something to keep
                // SkipWhitespace from spinning
                look = (char)0;
            }
        }

        /**
         * @return <c>true</c> if the specified name is a valid cell reference
         */
        private bool IsValidCellReference(String str) => IsValidCellReference(str.AsSpan());

        private bool IsValidCellReference(ReadOnlySpan<char> str)
        {
            //check range bounds against grid max
            bool result = CellReference.ClassifyCellReference(str, _ssVersion) == NameType.Cell;
            if (result)
            {
                /*
                 * Check if the argument is a function. Certain names can be either a cell reference or a function name
                 * depending on the contenxt. Compare the following examples in Excel 2007:
                 * (a) LOG10(100) + 1
                 * (b) LOG10 + 1
                 * In (a) LOG10 is a name of a built-in function. In (b) LOG10 is a cell reference
                 */
                bool isFunc = FunctionMetadataRegistry.GetFunctionByName(str.ToString().ToUpper()) != null;
                if (isFunc)
                {
                    int savePointer = _pointer;
                    ResetPointer(_pointer + str.Length);
                    SkipWhite();
                    // open bracket indicates that the argument is a function,
                    // the returning value should be false, i.e. "not a valid cell reference"
                    result = look != '(';
                    ResetPointer(savePointer);
                }
            }
            return result;
        }


        /**
         * Note - Excel Function names are 'case aware but not case sensitive'.  This method may end
         * up creating a defined name record in the workbook if the specified name is not an internal
         * Excel Function, and Has not been encountered before.
         * 
         * Side effect: creates workbook name if name is not recognized (name is probably a UDF)
         *
         * @param name case preserved Function name (as it was entered/appeared in the formula).
         */
        private ParseNode Function(String name)
        {
            Ptg nameToken = null;
            if (!AbstractFunctionPtg.IsBuiltInFunctionName(name))
            {
                // user defined Function
                // in the Token tree, the name is more or less the first argument

                if (_book == null)
                {
                    // Only test cases omit the book (expecting it not to be needed)
                    throw new InvalidOperationException("Need book to evaluate name '" + name + "'");
                }

                // Check to see if name is a named range in the workbook
                IEvaluationName hName = _book.GetName(name, _sheetIndex);
                if (hName != null)
                {
                    if (!hName.IsFunctionName)
                    {
                        throw new FormulaParseException("Attempt to use name '" + name
                                + "' as a function, but defined name in workbook does not refer to a function");
                    }

                    // calls to user-defined functions within the workbook
                    // get a Name token which points to a defined name record
                    nameToken = hName.CreatePtg();
                }
                else
                {
                    // Check if name is an external names table
                    nameToken = _book.GetNameXPtg(name, null);
                    if (nameToken == null)
                    {
                        // name is not an internal or external name
                        //if (log.check(POILogger.WARN))
                        //{
                        //    log.log(POILogger.WARN,
                        //            "FormulaParser.function: Name '" + name + "' is completely unknown in the current workbook.");
                        //}
                        // name is probably the name of an unregistered User-Defined Function
                        switch (_book.GetSpreadsheetVersion().Name)
                        {
                            case  "EXCEL97":
                                // HSSFWorkbooks require a name to be added to Workbook defined names table
                                AddName(name);
                                hName = _book.GetName(name, _sheetIndex);
                                nameToken = hName.CreatePtg();
                                break;
                            case "EXCEL2007":
                                // XSSFWorkbooks store formula names as strings.
                                nameToken = new NameXPxg(name);
                                break;
                            default:
                                throw new Exception("Unexpected spreadsheet version: " + _book.GetSpreadsheetVersion().Name);
                        }
                    }
                }
            }

            Match('(');
            ParseNode[] args = Arguments();
            Match(')');

            return GetFunction(name, nameToken, args);
        }
        /**
	     * Adds a name (named range or user defined function) to underlying workbook's names table
	     * @param functionName
	     */
        private void AddName(String functionName)
        {
            IName name = _book.CreateName();
            name.SetFunction(true);
            name.NameName = (functionName);
            name.SheetIndex = (_sheetIndex);
        }
        /**
         * Generates the variable Function ptg for the formula.
         * 
         * For IF Formulas, Additional PTGs are Added To the Tokens
     * @param name a {@link NamePtg} or {@link NameXPtg} or <code>null</code>
         * @return Ptg a null is returned if we're in an IF formula, it needs extreme manipulation and is handled in this Function
         */
        private ParseNode GetFunction(String name, Ptg namePtg, ParseNode[] args)
        {

            FunctionMetadata fm = FunctionMetadataRegistry.GetFunctionByName(name.ToUpper());
            int numArgs = args.Length;
            if (fm == null)
            {
                if (namePtg == null)
                {
                    throw new InvalidOperationException("NamePtg must be supplied for external Functions");
                }
                // must be external Function
                ParseNode[] allArgs = new ParseNode[numArgs + 1];
                allArgs[0] = new ParseNode(namePtg);
                System.Array.Copy(args, 0, allArgs, 1, numArgs);
                return new ParseNode(FuncVarPtg.Create(name, (byte)(numArgs + 1)), allArgs);
            }

            if (namePtg != null)
            {
                throw new InvalidOperationException("NamePtg no applicable To internal Functions");
            }
            bool IsVarArgs = !fm.HasFixedArgsLength;
            int funcIx = fm.Index;
        if (funcIx == FunctionMetadataRegistry.FUNCTION_INDEX_SUM && args.Length == 1) {
            // Excel encodes the sum of a single argument as tAttrSum
            // POI does the same for consistency, but this is not critical
            return new ParseNode(AttrPtg.GetSumSingle(), args);
            // The code below would encode tFuncVar(SUM) which seems to do no harm
        }
            ValidateNumArgs(args.Length, fm);

            AbstractFunctionPtg retval;
            if (IsVarArgs)
            {
                retval = FuncVarPtg.Create(name, (byte)numArgs);
            }
            else
            {
                retval = FuncPtg.Create(funcIx);
            }
            return new ParseNode(retval, args);
        }

        private void ValidateNumArgs(int numArgs, FunctionMetadata fm)
        {
            if (numArgs < fm.MinParams)
            {
                String msg = "Too few arguments to function '" + fm.Name + "'. ";
                if (fm.HasFixedArgsLength)
                {
                    msg += "Expected " + fm.MinParams;
                }
                else
                {
                    msg += "At least " + fm.MinParams + " were expected";
                }
                msg += " but got " + numArgs + ".";
                throw new FormulaParseException(msg);
            }
            //the maximum number of arguments depends on the Excel version
            int maxArgs;
            if (fm.HasUnlimitedVarags)
            {
                if (_book != null)
                {
                    maxArgs = _book.GetSpreadsheetVersion().MaxFunctionArgs;
                }
                else
                {
                    //_book can be omitted by test cases
                    maxArgs = fm.MaxParams; // just use BIFF8
                }
            }
            else
            {
                maxArgs = fm.MaxParams;
            }
            if (numArgs > maxArgs)
            {
                String msg = "Too many arguments to function '" + fm.Name + "'. ";
                if (fm.HasFixedArgsLength)
                {
                    msg += "Expected " + fm.MaxParams;
                }
                else
                {
                    msg += "At most " + fm.MaxParams + " were expected";
                }
                msg += " but got " + numArgs + ".";
                throw new FormulaParseException(msg);
            }
        }

        private static bool IsArgumentDelimiter(char ch)
        {
            return ch == ',' || ch == ')';
        }

        /** Get arguments To a Function */
        private ParseNode[] Arguments()
        {
            //average 2 args per Function
            ArrayList temp = new ArrayList(2);
            SkipWhite();
            if (look == ')')
            {
                return ParseNode.EMPTY_ARRAY;
            }

            bool missedPrevArg = true;
            int numArgs = 0;
            while (true)
            {
                SkipWhite();
                if (IsArgumentDelimiter(look))
                {
                    if (missedPrevArg)
                    {
                        temp.Add(new ParseNode(MissingArgPtg.instance));
                        numArgs++;
                    }
                    if (look == ')')
                    {
                        break;
                    }
                    Match(',');
                    missedPrevArg = true;
                    continue;
                }
                temp.Add(ComparisonExpression());
                numArgs++;
                missedPrevArg = false;
                SkipWhite();
                if (!IsArgumentDelimiter(look))
                {
                    throw expected("',' or ')'");
                }
            }
            ParseNode[] result = (ParseNode[])temp.ToArray(typeof(ParseNode));
            return result;
        }

        /** Parse and Translate a Math Factor  */
        private ParseNode PowerFactor()
        {
            ParseNode result = PercentFactor();
            while (true)
            {
                SkipWhite();
                if (look != '^')
                {
                    return result;
                }
                Match('^');
                ParseNode other = PercentFactor();
                result = new ParseNode(PowerPtg.instance, result, other);
            }
        }

        private ParseNode PercentFactor()
        {
            ParseNode result = ParseSimpleFactor();
            while (true)
            {
                SkipWhite();
                if (look != '%')
                {
                    return result;
                }
                Match('%');
                result = new ParseNode(PercentPtg.instance, result);
            }
        }



        /**
         * factors (without ^ or % )
         */
        private ParseNode ParseSimpleFactor()
        {
            SkipWhite();
            switch (look)
            {
                case '#':
                    return new ParseNode(ErrPtg.ValueOf(ParseErrorLiteral()));
                case '-':
                    Match('-');
                    return ParseUnary(false);
                case '+':
                    Match('+');
                    return ParseUnary(true);
                case '(':
                    Match('(');
                    ParseNode inside = UnionExpression();
                    Match(')');
                    return new ParseNode(ParenthesisPtg.instance, inside);
                case '"':
                    return new ParseNode(new StringPtg(ParseStringLiteral()));
                case '{':
                    Match('{');
                    ParseNode arrayNode = ParseArray();
                    Match('}');
                    return arrayNode;
            }
            // named ranges and tables can start with underscore or backslash
            // see https://support.office.com/en-us/article/Define-and-use-names-in-formulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64?ui=en-US&rs=en-US&ad=US#bmsyntax_rules_for_names
            if (IsAlpha(look) || Char.IsDigit(look) || look == '\'' || look == '[' || look == '_' || look == '\\')
            {
                return ParseRangeExpression();
            }
            if (look == '.')
            {
                return new ParseNode(ParseNumber());
            }
            throw expected("cell ref or constant literal");
        }
        private ParseNode ParseUnary(bool isPlus)
        {

            bool numberFollows = IsDigit(look) || look == '.';
            ParseNode factor = PowerFactor();

            if (numberFollows)
            {
                // + or - directly next to a number is parsed with the number

                Ptg token = factor.GetToken();
                if (token is NumberPtg)
                {
                    if (isPlus)
                    {
                        return factor;
                    }
                    token = new NumberPtg(-((NumberPtg)token).Value);
                    return new ParseNode(token);
                }
                if (token is IntPtg)
                {
                    if (isPlus)
                    {
                        return factor;
                    }
                    int intVal = ((IntPtg)token).Value;
                    // note - cannot use IntPtg for negatives
                    token = new NumberPtg(-intVal);
                    return new ParseNode(token);
                }
            }
            return new ParseNode(isPlus ? UnaryPlusPtg.instance : UnaryMinusPtg.instance, factor);
        }

        private ParseNode ParseArray()
        {
            List<Object[]> rowsData = new List<Object[]>();
            while (true)
            {
                Object[] singleRowData = ParseArrayRow();
                rowsData.Add(singleRowData);
                if (look == '}')
                {
                    break;
                }
                if (look != ';')
                {
                    throw expected("'}' or ';'");
                }
                Match(';');
            }
            int nRows = rowsData.Count;
            Object[][] values2d = new Object[nRows][];
            values2d = (Object[][])rowsData.ToArray();
            int nColumns = values2d[0].Length;
            CheckRowLengths(values2d, nColumns);

            return new ParseNode(new ArrayPtg(values2d));
        }
        private void CheckRowLengths(Object[][] values2d, int nColumns)
        {
            for (int i = 0; i < values2d.Length; i++)
            {
                int rowLen = values2d[i].Length;
                if (rowLen != nColumns)
                {
                    throw new FormulaParseException("Array row " + i + " Has length " + rowLen
                            + " but row 0 Has length " + nColumns);
                }
            }
        }

        private Object[] ParseArrayRow()
        {
            ArrayList temp = new ArrayList();
            while (true)
            {
                temp.Add(ParseArrayItem());
                SkipWhite();
                switch (look)
                {
                    case '}':
                    case ';':
                        break;
                    case ',':
                        Match(',');
                        continue;
                    default:
                        throw expected("'}' or ','");

                }
                break;
            }

            Object[] result = new Object[temp.Count];
            result = temp.ToArray();
            return result;
        }

        private Object ParseArrayItem()
        {
            SkipWhite();
            switch (look)
            {
                case '"': return ParseStringLiteral();
                case '#': return ErrorConstant.ValueOf(ParseErrorLiteral());
                case 'F':
                case 'f':
                case 'T':
                case 't':
                    return ParseBooleanLiteral();
                case '-':
                    Match('-');
                    SkipWhite();
                    return ConvertArrayNumber(ParseNumber(), false);
            }
            // else assume number
            return ConvertArrayNumber(ParseNumber(), true);
        }

        private Boolean ParseBooleanLiteral()
        {
            String iden = ParseUnquotedIdentifier();
            if ("TRUE".Equals(iden, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if ("FALSE".Equals(iden, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
            throw expected("'TRUE' or 'FALSE'");
        }

        private static Double ConvertArrayNumber(Ptg ptg, bool isPositive)
        {
            double value;
            if (ptg is IntPtg)
            {
                value = ((IntPtg)ptg).Value;
            }
            else if (ptg is NumberPtg)
            {
                value = ((NumberPtg)ptg).Value;
            }
            else
            {
                throw new Exception("Unexpected ptg (" + ptg.GetType().Name + ")");
            }
            if (!isPositive)
            {
                value = -value;
            }
            return value;
        }

        private Ptg ParseNumber()
        {
            String number2 = null;
            String exponent = null;
            String number1 = GetNum();

            if (look == '.')
            {
                GetChar();
                number2 = GetNum();
            }

            if (look == 'E')
            {
                GetChar();

                String sign = "";
                if (look == '+')
                {
                    GetChar();
                }
                else if (look == '-')
                {
                    GetChar();
                    sign = "-";
                }

                String number = GetNum();
                if (number == null)
                {
                    throw expected("int");
                }
                exponent = sign + number;
            }

            if (number1 == null && number2 == null)
            {
                throw expected("int");
            }

            return GetNumberPtgFromString(number1, number2, exponent);
        }


        private int ParseErrorLiteral()
        {
            Match('#');
            String part1 = ParseUnquotedIdentifier().ToUpper();

            switch (part1[0])
            {
                case 'V':
                    {
                        FormulaError fe = FormulaError.VALUE;
                        if (part1.Equals(fe.Name))
                        {
                            Match('!');
                            return fe.Code;
                        }
                        throw expected(fe.String);
                    }
                case 'R':
                    {
                        FormulaError fe = FormulaError.REF;
                        if (part1.Equals(fe.Name))
                        {
                            Match('!');
                            return fe.Code;
                        }
                        throw expected(fe.String);
                    }
                case 'D':
                    {
                        FormulaError fe = FormulaError.DIV0;
                        if (part1.Equals("DIV"))
                        {
                            Match('/');
                            Match('0');
                            Match('!');
                            return fe.Code;
                        }
                        throw expected(fe.String);
                    }
                case 'N':
                    {
                        FormulaError fe = FormulaError.NAME;
                        if (part1.Equals(fe.Name))
                        {
                            // only one that ends in '?'
                            Match('?');
                            return fe.Code;
                        }
                        fe = FormulaError.NUM;
                        if (part1.Equals(fe.Name))
                        {
                            Match('!');
                            return fe.Code;
                        }
                        fe = FormulaError.NULL;
                        if (part1.Equals(fe.Name))
                        {
                            Match('!');
                            return fe.Code;
                        }
                        fe = FormulaError.NA;
                        if (part1.Equals("N"))
                        {
                            Match('/');
                            if (look != 'A' && look != 'a')
                            {
                                throw expected(fe.String);
                            }
                            Match(look);
                            // Note - no '!' or '?' suffix
                            return fe.Code;
                        }
                        throw expected("#NAME?, #NUM!, #NULL! or #N/A");
                    }
            }
            throw expected("#VALUE!, #REF!, #DIV/0!, #NAME?, #NUM!, #NULL! or #N/A");
        }


        /**
         * Get a PTG for an integer from its string representation.
         * return Int or Number Ptg based on size of input
         */
        private static Ptg GetNumberPtgFromString(String number1, String number2, String exponent)
        {
            StringBuilder number = new StringBuilder();

            if (number2 == null)
            {
                number.Append(number1);

                if (exponent != null)
                {
                    number.Append('E');
                    number.Append(exponent);
                }

                String numberStr = number.ToString();
                int intVal;
                try
                {
                    intVal = int.Parse(numberStr, CultureInfo.InvariantCulture);
                }
                catch (FormatException)
                {
                    return new NumberPtg(numberStr);
                }
                catch (OverflowException)
                {
                    return new NumberPtg(numberStr);
                }
                if (IntPtg.IsInRange(intVal))
                {
                    return new IntPtg(intVal);
                }
                return new NumberPtg(numberStr);
            }

            if (number1 != null)
            {
                number.Append(number1);
            }

            number.Append('.');
            number.Append(number2);

            if (exponent != null)
            {
                number.Append('E');
                number.Append(exponent);
            }

            return new NumberPtg(number.ToString());
        }


        private String ParseStringLiteral()
        {
            Match('"');

            StringBuilder Token = new StringBuilder();
            while (true)
            {
                if (look == '"')
                {
                    GetChar();
                    if (look != '"')
                    {
                        break;
                    }
                }
                Token.Append(look);
                GetChar();
            }
            return Token.ToString();
        }

        /** Parse and Translate a Math Term */
        private ParseNode Term()
        {
            ParseNode result = PowerFactor();
            while (true)
            {
                SkipWhite();
                Ptg operator1;
                switch (look)
                {
                    case '*':
                        Match('*');
                        operator1 = MultiplyPtg.instance;
                        break;
                    case '/':
                        Match('/');
                        operator1 = DividePtg.instance;
                        break;
                    default:
                        return result; // finished with Term
                }
                ParseNode other = PowerFactor();
                result = new ParseNode(operator1, result, other);
            }
        }

        private ParseNode ComparisonExpression()
        {
            ParseNode result = ConcatExpression();
            while (true)
            {
                SkipWhite();
                switch (look)
                {
                    case '=':
                    case '>':
                    case '<':
                        Ptg comparisonToken = GetComparisonToken();
                        ParseNode other = ConcatExpression();
                        result = new ParseNode(comparisonToken, result, other);
                        continue;
                }
                return result; // finished with predicate expression
            }
        }

        private Ptg GetComparisonToken()
        {
            if (look == '=')
            {
                Match(look);
                return EqualPtg.instance;
            }
            bool IsGreater = look == '>';
            Match(look);
            if (IsGreater)
            {
                if (look == '=')
                {
                    Match('=');
                    return GreaterEqualPtg.instance;
                }
                return GreaterThanPtg.instance;
            }
            switch (look)
            {
                case '=':
                    Match('=');
                    return LessEqualPtg.instance;
                case '>':
                    Match('>');
                    return NotEqualPtg.instance;
            }
            return LessThanPtg.instance;
        }


        private ParseNode ConcatExpression()
        {
            ParseNode result = AdditiveExpression();
            while (true)
            {
                SkipWhite();
                if (look != '&')
                {
                    break; // finished with concat expression
                }
                Match('&');
                ParseNode other = AdditiveExpression();
                result = new ParseNode(ConcatPtg.instance, result, other);
            }
            return result;
        }


        /** Parse and Translate an Expression */
        private ParseNode AdditiveExpression()
        {
            ParseNode result = Term();
            while (true)
            {
                SkipWhite();
                Ptg operator1;
                switch (look)
                {
                    case '+':
                        Match('+');
                        operator1 = AddPtg.instance;
                        break;
                    case '-':
                        Match('-');
                        operator1 = SubtractPtg.instance;
                        break;
                    default:
                        return result; // finished with Additive expression
                }
                ParseNode other = Term();
                result = new ParseNode(operator1, result, other);
            }
        }

        //{--------------------------------------------------------------}
        //{ Parse and Translate an Assignment Statement }
        /*
    procedure Assignment;
    var Name: string[8];
    begin
       Name := GetName;
       Match('=');
       Expression;

    end;
         **/


        /**
         *  API call To execute the parsing of the formula
         * 
         */
        private void Parse()
        {
            _pointer = 0;
            GetChar();
            _rootNode = UnionExpression();

            if (_pointer <= _formulaLength)
            {
                String msg = "Unused input [" + _formulaString.Substring(_pointer - 1)
                    + "] after attempting to parse the formula [" + _formulaString + "]";
                throw new FormulaParseException(msg);
            }
        }
        private ParseNode UnionExpression()
        {
            ParseNode result = IntersectionExpression();
            bool hasUnions = false;
            while (true)
            {
                SkipWhite();
                switch (look)
                {
                    case ',':
                        GetChar();
                        hasUnions = true;
                        ParseNode other = IntersectionExpression();
                        result = new ParseNode(UnionPtg.instance, result, other);
                        continue;
                }
                if (hasUnions)
                {
                    return AugmentWithMemPtg(result);
                }
                return result;
            }
        }
        private ParseNode IntersectionExpression()
        {
            ParseNode result = ComparisonExpression();
            bool hasIntersections = false;
            while (true)
            {
                SkipWhite();
                if (_inIntersection)
                {
                    int savePointer = _pointer;
                    // Don't getChar() as the space has already been eaten and recorded by SkipWhite().
                    try
                    {
                        ParseNode other = ComparisonExpression();
                        result = new ParseNode(IntersectionPtg.instance, result, other);
                        hasIntersections = true;
                        continue;
                    }
                    catch (FormulaParseException)
                    {
                        // if parsing for intersection fails we assume that we actually had an arbitrary
                        // whitespace and thus should simply skip this whitespace
                        ResetPointer(savePointer);
                    }
                }
                if (hasIntersections)
                {
                    return AugmentWithMemPtg(result);
                }
                return result;
            }
        }

        private Ptg[] GetRPNPtg(FormulaType formulaType)
        {
            OperandClassTransformer oct = new OperandClassTransformer(formulaType);
            // RVA is for 'operand class': 'reference', 'value', 'array'
            oct.TransformFormula(_rootNode);
            return ParseNode.ToTokenArray(_rootNode);
        }
    }
}