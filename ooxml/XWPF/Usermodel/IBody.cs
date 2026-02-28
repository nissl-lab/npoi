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

namespace NPOI.XWPF.UserModel
{
    using System;
    using System.Collections.Generic;
    using System.Xml;
    using NPOI.OpenXmlFormats.Wordprocessing;

    /**
     * An IBody represents the different parts of the document which
     * can contain collections of Paragraphs and Tables. It provides a
     * common way to work with these and their contents.
     * Typically, this is something like a XWPFDocument, or one of
     *  the parts in it like XWPFHeader, XWPFFooter, XWPFTableCell
     *
     */
    public interface IBody
    {
        /**
         * returns the Part, to which the body belongs, which you need for Adding relationship to other parts
         * Actually it is needed of the class XWPFTableCell. Because you have to know to which part the tableCell
         * belongs.
         * @return the Part, to which the body belongs
         */
        POIXMLDocumentPart Part
        {
            get;
        }

        /**
         * Get the PartType of the body, for example
         * DOCUMENT, HEADER, FOOTER,	FOOTNOTE, 
         * @return the PartType of the body
         */
        BodyType PartType { get; }

        /**
         * Returns an Iterator with paragraphs and tables, 
         *  in the order that they occur in the text.
         */
        IList<IBodyElement> BodyElements { get; }

        /**
         * Returns the paragraph(s) that holds
         *  the text of the header or footer.
         */
        IList<XWPFParagraph> Paragraphs { get; }

        /**
         * Return the table(s) that holds the text
         *  of the IBodyPart, for complex cases
         *  where a paragraph isn't used.
         */
        IList<XWPFTable> Tables { get; }

        /**
         * if there is a corresponding {@link XWPFParagraph} of the parameter ctTable in the paragraphList of this header or footer
         * the method will return this paragraph
         * if there is no corresponding {@link XWPFParagraph} the method will return null 
         * @param p is instance of CTP and is searching for an XWPFParagraph
         * @return null if there is no XWPFParagraph with an corresponding CTPparagraph in the paragraphList of this header or footer
         * 		   XWPFParagraph with the correspondig CTP p
         */
        XWPFParagraph GetParagraph(CT_P p);

        /**
         * if there is a corresponding {@link XWPFTable} of the parameter ctTable in the tableList of this header
         * the method will return this table
         * if there is no corresponding {@link XWPFTable} the method will return null 
         * @param ctTable
         */
        XWPFTable GetTable(CT_Tbl ctTable);

        /**
         * Returns the paragraph that of position pos
         */
        XWPFParagraph GetParagraphArray(int pos);

        /**
         * Returns the table at position pos
         */
        XWPFTable GetTableArray(int pos);

        /**
         *inserts a new paragraph at position of the cursor
         * @param cursor
         */
        XWPFParagraph InsertNewParagraph(XmlDocument cursor);

        /**
         * inserts a new Table at the cursor position.
         * @param cursor
         */
        XWPFTable InsertNewTbl(/*XmlCursor*/XmlDocument cursor);

        /**
         * inserts a new Table at position pos
         * @param pos
         * @param table
         */
        void InsertTable(int pos, XWPFTable table);

        /**
         * returns the TableCell to which the Table belongs
         * @param cell
         */
        XWPFTableCell GetTableCell(CT_Tc cell);

        /**
         * Return XWPFDocument
         */
        XWPFDocument GetXWPFDocument();

    }


}