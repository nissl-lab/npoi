/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HWPF.UserModel;
using NPOI.Util;
using System.Xml;
using NPOI.HWPF.Model;
using System.Diagnostics;
using System.Text.RegularExpressions;
using NPOI.POIFS.FileSystem;


namespace NPOI.HWPF.Converter
{
    public abstract class AbstractWordConverter
    {
        private class Structure : IComparable<Structure>
        {
            internal int End;
            internal int Start;
            internal Object StructureObject;

            public Structure(Bookmark bookmark)
            {
                this.Start = bookmark.Start;
                this.End = bookmark.End;
                this.StructureObject = bookmark;
            }

            public Structure(Field field)
            {
                this.Start = field.GetFieldStartOffset();
                this.End = field.GetFieldEndOffset();
                this.StructureObject = field;
            }

            public override String ToString()
            {
                return "Structure [" + Start + "; " + End + "]: "
                        + StructureObject.ToString();
            }

            #region IComparable<Structure> 成员

            public int CompareTo(Structure other)
            {
                return Start < other.Start ? -1 : Start == other.Start ? 0 : 1;
            }

            #endregion
        }

        private static byte BEL_MARK = 7;

        private static byte FIELD_BEGIN_MARK = 19;

        private static byte FIELD_END_MARK = 21;

        private static byte FIELD_SEPARATOR_MARK = 20;

        private static POILogger logger = POILogFactory.GetLogger(typeof(AbstractWordConverter));

        private static byte SPECCHAR_AUTONUMBERED_FOOTNOTE_REFERENCE = 2;

        private static byte SPECCHAR_DRAWN_OBJECT = 8;

        protected static char UNICODECHAR_NONBREAKING_HYPHEN = '\u2011';

        protected static char UNICODECHAR_ZERO_WIDTH_SPACE = '\u200b';

        private static void AddToStructures(IList<Structure> structures, Structure structure)
        {
            IList<Structure> toRemove = new List<Structure>();
            foreach (Structure another in structures)
            {

                if (another.Start <= structure.Start
                        && another.End >= structure.Start)
                {
                    return;
                }

                if ((structure.Start < another.Start && another.Start < structure.End)
                        || (structure.Start < another.Start && another.End <= structure.End)
                        || (structure.Start <= another.Start && another.End < structure.End))
                {
                    //iterator.remove();
                    toRemove.Add(another);
                    continue;
                }
            }

            foreach (Structure s in toRemove)
                structures.Remove(s);
            structures.Add(structure);
        }
        private List<Bookmark> bookmarkStack = new List<Bookmark>();

        private FontReplacer fontReplacer = new DefaultFontReplacer();

        private PicturesManager picturesManager;

        /**
         * Special actions that need to be called after processing complete, like
         * updating stylesheets or building document notes list. Usually they are
         * called once, but it's okay to call them several times.
         */
        protected virtual void AfterProcess()
        {
            // by default no such actions needed
        }
        protected Triplet GetCharacterRunTriplet(CharacterRun characterRun)
        {
            Triplet original = new Triplet();
            original.bold = characterRun.IsBold();
            original.italic = characterRun.IsItalic();
            original.fontName = characterRun.GetFontName();
            Triplet updated = GetFontReplacer().Update(original);
            return updated;
        }
        public abstract XmlDocument Document { get; }
        public FontReplacer GetFontReplacer()
        {
            return fontReplacer;
        }
        protected int GetNumberColumnsSpanned(int[] tableCellEdges,
            int currentEdgeIndex, TableCell tableCell)
        {
            int nextEdgeIndex = currentEdgeIndex;
            int colSpan = 0;
            int cellRightEdge = tableCell.GetLeftEdge() + tableCell.GetWidth();
            while (tableCellEdges[nextEdgeIndex] < cellRightEdge)
            {
                colSpan++;
                nextEdgeIndex++;
            }
            return colSpan;
        }
        protected int GetNumberRowsSpanned(Table table, int currentRowIndex,
            int currentColumnIndex, TableCell tableCell)
        {
            if (!tableCell.IsFirstVerticallyMerged())
                return 1;

            int numRows = table.NumRows;

            int count = 1;
            for (int r1 = currentRowIndex + 1; r1 < numRows; r1++)
            {
                TableRow nextRow = table.GetRow(r1);
                if (currentColumnIndex >= nextRow.NumCells())
                    break;
                TableCell nextCell = nextRow.GetCell(currentColumnIndex);
                if (!nextCell.IsVerticallyMerged()
                        || nextCell.IsFirstVerticallyMerged())
                    break;
                count++;
            }
            return count;
        }

        public PicturesManager GetPicturesManager()
        {
            return picturesManager;
        }

        protected int getTableCellEdgesIndexSkipCount(Table table, int r,
            int[] tableCellEdges, int currentEdgeIndex, int c, TableCell tableCell)
        {
            TableCell upperCell = null;
            for (int r1 = r - 1; r1 >= 0; r1--)
            {
                TableRow row = table.GetRow(r1);
                if (row == null || c >= row.NumCells())
                    continue;

                TableCell prevCell = row.GetCell(c);
                if (prevCell != null && prevCell.IsFirstVerticallyMerged())
                {
                    upperCell = prevCell;
                    break;
                }
            }
            if (upperCell == null)
            {
                logger.Log(POILogger.WARN, "First vertically merged cell for ",
                        tableCell, " not found");
                return 0;
            }

            return GetNumberColumnsSpanned(tableCellEdges, currentEdgeIndex,
                    tableCell);
        }
        protected abstract void OutputCharacters(XmlElement block,
            CharacterRun characterRun, String text);

        /**
         * Wrap range into bookmark(s) and process it. All bookmarks have starts
         * equal to range start and ends equal to range end. Usually it's only one
         * bookmark.
         */
        protected abstract void ProcessBookmarks(HWPFDocumentCore wordDocument,
                XmlElement currentBlock, Range range, int currentTableLevel,
                IList<Bookmark> rangeBookmarks);

        protected bool ProcessCharacters(HWPFDocumentCore wordDocument,  int currentTableLevel, Range range, XmlElement block)
        {
            if (range == null)
                return false;

            bool haveAnyText = false;

            /*
             * In text there can be fields, bookmarks, may be other structures (code
             * below allows extension). Those structures can overlaps, so either we
             * should process char-by-char (slow) or find a correct way to
             * reconstruct the structure of range -- sergey
             */
            IList<Structure> structures = new List<Structure>();
            if (wordDocument is HWPFDocument)
            {
                HWPFDocument doc = (HWPFDocument)wordDocument;

                Dictionary<int, List<Bookmark>> rangeBookmarks = doc.GetBookmarks()
                        .GetBookmarksStartedBetween(range.StartOffset, range.EndOffset);

                if (rangeBookmarks != null)
                {
                    foreach (KeyValuePair<int, List<Bookmark>> kv in rangeBookmarks)
                    {
                        List<Bookmark> lists = kv.Value;
                        foreach (Bookmark bookmark in lists)
                        {
                            if (!bookmarkStack.Contains(bookmark))
                                AddToStructures(structures, new Structure(bookmark));
                        }
                    }
                }

                // TODO: dead fields?
                for (int c = 0; c < range.NumCharacterRuns; c++)
                {
                    CharacterRun characterRun = range.GetCharacterRun(c);
                    if (characterRun == null)
                        throw new NullReferenceException();
                    Field aliveField = ((HWPFDocument)wordDocument).GetFields()
                            .GetFieldByStartOffset(FieldsDocumentPart.MAIN,
                                    characterRun.StartOffset);
                    if (aliveField != null)
                    {
                        AddToStructures(structures, new Structure(aliveField));
                    }
                }
            }

            //structures = new ArrayList<Structure>( structures );
            //Collections.sort( structures );
            SortedList<Structure, Structure> sl = new SortedList<Structure, Structure>();
            foreach (Structure s in structures)
                sl.Add(s, s);
            structures.Clear();
            ((List<Structure>)structures).AddRange(sl.Values);

            int previous = range.StartOffset;
            foreach (Structure structure in structures)
            {
                if (structure.Start != previous)
                {
                    Range subrange = new Range(previous, structure.Start, range);
                    //{
                    //    public String toString()
                    //    {
                    //        return "BetweenStructuresSubrange " + super.ToString();
                    //    }
                    //};
                    ProcessCharacters(wordDocument, currentTableLevel, subrange, block);
                }

                if (structure.StructureObject is Bookmark)
                {
                    // other bookmarks with same boundaries
                    IList<Bookmark> bookmarks = new List<Bookmark>();
                    IEnumerator<List<Bookmark>> iterator = ((HWPFDocument)wordDocument).GetBookmarks().GetBookmarksStartedBetween(structure.Start, structure.Start + 1).Values.GetEnumerator();
                    iterator.MoveNext();
                    foreach (Bookmark bookmark in iterator.Current)
                    {
                        if (bookmark.Start == structure.Start
                                && bookmark.End == structure.End)
                        {
                            bookmarks.Add(bookmark);
                        }
                    }

                    bookmarkStack.AddRange(bookmarks);
                    try
                    {
                        int end = Math.Min(range.EndOffset, structure.End);
                        Range subrange = new Range(structure.Start, end, range);
                        /*{
                            public String toString()
                            {
                                return "BookmarksSubrange " + super.ToString();
                            }
                        };*/

                        ProcessBookmarks(wordDocument, block, subrange,
                                currentTableLevel, bookmarks);
                    }
                    finally
                    {
                        bookmarkStack.RemoveAll((e) => { return bookmarks.Contains(e); });
                    }
                }
                else if (structure.StructureObject is Field)
                {
                    Field field = (Field)structure.StructureObject;
                    ProcessField((HWPFDocument)wordDocument, range, currentTableLevel, field, block);
                }
                else
                {
                    throw new NotSupportedException("NYI: " + structure.StructureObject.GetType().ToString());
                }

                previous = Math.Min(range.EndOffset, structure.End);
            }

            if (previous != range.StartOffset)
            {
                if (previous > range.EndOffset)
                {
                    logger.Log(POILogger.WARN, "Latest structure in ", range,
                            " ended at #" + previous, " after range boundaries [",
                            range.StartOffset + "; " + range.EndOffset,
                            ")");
                    return true;
                }

                if (previous < range.EndOffset)
                {
                    Range subrange = new Range(previous, range.EndOffset, range);
                    /*{
                        @Override
                        public String toString()
                        {
                            return "AfterStructureSubrange " + super.ToString();
                        }
                    };*/
                    ProcessCharacters(wordDocument, currentTableLevel, subrange,
                            block);
                }
                return true;
            }

            for (int c = 0; c < range.NumCharacterRuns; c++)
            {
                CharacterRun characterRun = range.GetCharacterRun(c);

                if (characterRun == null)
                    throw new NullReferenceException();

                if (wordDocument is HWPFDocument && ((HWPFDocument)wordDocument).GetPicturesTable().HasPicture(characterRun))
                {
                    HWPFDocument newFormat = (HWPFDocument)wordDocument;
                    Picture picture = newFormat.GetPicturesTable().ExtractPicture(characterRun, true);

                    ProcessImage(block, characterRun.Text[0] == 0x01, picture);
                    continue;
                }

                string text = characterRun.Text;
                byte[] textByte = System.Text.Encoding.GetEncoding("iso-8859-1").GetBytes(text);
                //if ( text.getBytes().length == 0 )
                if (textByte.Length == 0)
                    continue;

                if (characterRun.IsSpecialCharacter())
                {
                    if (text[0] == SPECCHAR_AUTONUMBERED_FOOTNOTE_REFERENCE && (wordDocument is HWPFDocument))
                    {
                        HWPFDocument doc = (HWPFDocument)wordDocument;
                        ProcessNoteAnchor(doc, characterRun, block);
                        continue;
                    }
                    if (text[0] == SPECCHAR_DRAWN_OBJECT
                            && (wordDocument is HWPFDocument))
                    {
                        HWPFDocument doc = (HWPFDocument)wordDocument;
                        ProcessDrawnObject(doc, characterRun, block);
                        continue;
                    }
                    if (characterRun.IsOle2() && (wordDocument is HWPFDocument))
                    {
                        HWPFDocument doc = (HWPFDocument)wordDocument;
                        ProcessOle2(doc, characterRun, block);
                        continue;
                    }
                }
                if (textByte[0] == FIELD_BEGIN_MARK)
                //if ( text.getBytes()[0] == FIELD_BEGIN_MARK )
                {
                    if (wordDocument is HWPFDocument)
                    {
                        Field aliveField = ((HWPFDocument)wordDocument).GetFields().GetFieldByStartOffset(
                                        FieldsDocumentPart.MAIN, characterRun.StartOffset);
                        if (aliveField != null)
                        {
                            ProcessField(((HWPFDocument)wordDocument), range,
                                    currentTableLevel, aliveField, block);

                            int continueAfter = aliveField.GetFieldEndOffset();
                            while (c < range.NumCharacterRuns
                                    && range.GetCharacterRun(c).EndOffset <= continueAfter)
                                c++;

                            if (c < range.NumCharacterRuns)
                                c--;

                            continue;
                        }
                    }

                    int skipTo = TryDeadField(wordDocument, range,
                            currentTableLevel, c, block);

                    if (skipTo != c)
                    {
                        c = skipTo;
                        continue;
                    }

                    continue;
                }
                if (textByte[0] == FIELD_SEPARATOR_MARK)
                {
                    // shall not appear without FIELD_BEGIN_MARK
                    continue;
                }
                if (textByte[0] == FIELD_END_MARK)
                {
                    // shall not appear without FIELD_BEGIN_MARK
                    continue;
                }

                if (characterRun.IsSpecialCharacter() || characterRun.IsObj()
                        || characterRun.IsOle2())
                {
                    continue;
                }

                if (text.EndsWith("\r")
                        || (text[text.Length - 1] == BEL_MARK && currentTableLevel != int.MinValue))
                    text = text.Substring(0, text.Length - 1);

                {
                    // line breaks
                    StringBuilder stringBuilder = new StringBuilder();
                    foreach (char charChar in text.ToCharArray())
                    {
                        if (charChar == 11)
                        {
                            if (stringBuilder.Length > 0)
                            {
                                OutputCharacters(block, characterRun,
                                        stringBuilder.ToString());
                                stringBuilder.Length = 0;
                            }
                            ProcessLineBreak(block, characterRun);
                        }
                        else if (charChar == 30)
                        {
                            // Non-breaking hyphens are stored as ASCII 30
                            stringBuilder.Append(UNICODECHAR_NONBREAKING_HYPHEN);
                        }
                        else if (charChar == 31)
                        {
                            // Non-required hyphens to zero-width space
                            stringBuilder.Append(UNICODECHAR_ZERO_WIDTH_SPACE);
                        }
                        else if (charChar >= 0x20 || charChar == 0x09
                                || charChar == 0x0A || charChar == 0x0D)
                        {
                            stringBuilder.Append(charChar);
                        }
                    }
                    if (stringBuilder.Length > 0)
                    {
                        OutputCharacters(block, characterRun,
                                stringBuilder.ToString());
                        stringBuilder.Length = 0;
                    }
                }

                haveAnyText |= text.Trim().Length != 0;
            }

            return haveAnyText;
        }
        protected void ProcessDeadField(HWPFDocumentCore wordDocument,
                XmlElement currentBlock, Range range, int currentTableLevel,
                int beginMark, int separatorMark, int endMark)
        {
            StringBuilder debug = new StringBuilder("Unsupported field type: \n");
            for (int i = beginMark; i <= endMark; i++)
            {
                debug.Append("\t");
                debug.Append(range.GetCharacterRun(i));
                debug.Append("\n");
            }
            logger.Log(POILogger.WARN, debug);

            Range deadFieldValueSubrage = new Range(range.GetCharacterRun(
                    separatorMark).StartOffset + 1, range.GetCharacterRun(
                    endMark).StartOffset, range);
            //{
            //    @Override
            //    public String toString()
            //    {
            //        return "DeadFieldValueSubrange (" + super.ToString() + ")";
            //    }
            //};

            // just output field value
            if (separatorMark + 1 < endMark)
                ProcessCharacters(wordDocument, currentTableLevel,
                        deadFieldValueSubrage, currentBlock);

            return;
        }

        protected Field ProcessDeadField(HWPFDocumentCore wordDocument,
                Range charactersRange, int currentTableLevel, int startOffset,
                XmlElement currentBlock)
        {
            if (!(wordDocument is HWPFDocument))
                return null;

            HWPFDocument hwpfDocument = (HWPFDocument)wordDocument;
            Field field = hwpfDocument.GetFields().GetFieldByStartOffset(
                    FieldsDocumentPart.MAIN, startOffset);
            if (field == null)
                return null;

            ProcessField(hwpfDocument, charactersRange, currentTableLevel, field,
                    currentBlock);

            return field;
        }

        public void ProcessDocument(HWPFDocumentCore wordDocument)
        {
            try
            {
                NPOI.HPSF.SummaryInformation summaryInformation = wordDocument.SummaryInformation;
                if (summaryInformation != null)
                {
                    ProcessDocumentInformation(summaryInformation);
                }
            }
            catch (Exception exc)
            {
                logger.Log(POILogger.WARN, "Unable to process document summary information: ", exc, exc);
            }

            Range docRange = wordDocument.GetRange();

            if (docRange.NumSections == 1)
            {
                ProcessSingleSection(wordDocument, docRange.GetSection(0));
                AfterProcess();
                return;
            }

            ProcessDocumentPart(wordDocument, docRange);
            AfterProcess();
        }

        protected abstract void ProcessDocumentInformation(NPOI.HPSF.SummaryInformation summaryInformation);

        protected virtual void ProcessDocumentPart(HWPFDocumentCore wordDocument,
                 Range range)
        {
            for (int s = 0; s < range.NumSections; s++)
            {
                ProcessSection(wordDocument, range.GetSection(s), s);
            }
        }

        protected void ProcessDrawnObject(HWPFDocument doc,
                CharacterRun characterRun, XmlElement block)
        {
            if (GetPicturesManager() == null)
                return;
            // TODO: support headers
            OfficeDrawing officeDrawing = doc.GetOfficeDrawingsMain().GetOfficeDrawingAt(characterRun.StartOffset);
            if (officeDrawing == null)
            {
                logger.Log(POILogger.WARN, "Characters #" + characterRun
                        + " references missing drawn object");
                return;
            }

            byte[] pictureData = officeDrawing.GetPictureData();
            if (pictureData == null)
                // usual shape?
                return;

            PictureType type = PictureType.FindMatchingType(pictureData);
            String path = GetPicturesManager().SavePicture(pictureData, type,
                    "s" + characterRun.StartOffset + "." + type);

            ProcessDrawnObject(doc, characterRun, officeDrawing, path, block);
        }

        protected abstract void ProcessDrawnObject(HWPFDocument doc,
                CharacterRun characterRun, OfficeDrawing officeDrawing,
                String path, XmlElement block);

        protected abstract void ProcessEndnoteAutonumbered(
                HWPFDocument wordDocument, int noteIndex, XmlElement block,
                Range endnoteTextRange);

        protected void ProcessField(HWPFDocument wordDocument, Range parentRange,
                int currentTableLevel, Field field, XmlElement currentBlock)
        {
            switch (field.Type)
            {
                case 37: // page reference
                    {
                        Range firstSubrange = field.FirstSubrange(parentRange);
                        if (firstSubrange != null)
                        {
                            String formula = firstSubrange.Text;
                            Regex pagerefPattern = new Regex("[ \\t\\r\\n]*PAGEREF ([^ ]*)[ \\t\\r\\n]*\\\\h[ \\t\\r\\n]*");
                            Match match = pagerefPattern.Match(formula);
                            if (match.Success)
                            {
                                String pageref = match.Groups[1].Value;
                                ProcessPageref(wordDocument, currentBlock,
                                        field.SecondSubrange(parentRange),
                                        currentTableLevel, pageref);
                                return;
                            }
                            //Pattern pagerefPattern = Pattern
                            //        .compile("[ \\t\\r\\n]*PAGEREF ([^ ]*)[ \\t\\r\\n]*\\\\h[ \\t\\r\\n]*");
                            //Matcher matcher = pagerefPattern.matcher(formula);
                            //if (matcher.find())
                            //{
                            //    String pageref = matcher.group(1);
                            //    processPageref(wordDocument, currentBlock,
                            //            field.secondSubrange(parentRange),
                            //            currentTableLevel, pageref);
                            //    return;
                            //}
                        }
                        break;
                    }
                case 58: // Embedded Object
                    {
                        if (!field.HasSeparator())
                        {
                            logger.Log(POILogger.WARN, parentRange + " contains " + field
                                    + " with 'Embedded Object' but without separator mark");
                            return;
                        }

                        CharacterRun separator = field.GetMarkSeparatorCharacterRun(parentRange);

                        if (separator.IsOle2())
                        {
                            // the only supported so far
                            bool processed = ProcessOle2(wordDocument, separator,
                                    currentBlock);

                            // if we didn't output OLE - output field value
                            if (!processed)
                            {
                                ProcessCharacters(wordDocument, currentTableLevel,
                                        field.SecondSubrange(parentRange), currentBlock);
                            }

                            return;
                        }

                        break;
                    }
                case 88: // hyperlink
                    {
                        Range firstSubrange = field.FirstSubrange(parentRange);
                        if (firstSubrange != null)
                        {
                            String formula = firstSubrange.Text;
                            Regex hyperlinkPattern = new Regex("[ \\t\\r\\n]*HYPERLINK \"(.*)\"[ \\t\\r\\n]*");
                            Match match = hyperlinkPattern.Match(formula);
                            if (match.Success)
                            {
                                String hyperlink = match.Groups[1].Value;
                                ProcessHyperlink(wordDocument, currentBlock,
                                        field.SecondSubrange(parentRange),
                                        currentTableLevel, hyperlink);
                                return;
                            }
                            //Pattern hyperlinkPattern = Pattern
                            //        .compile("[ \\t\\r\\n]*HYPERLINK \"(.*)\"[ \\t\\r\\n]*");
                            //Matcher matcher = hyperlinkPattern.matcher(formula);
                            //if (matcher.find())
                            //{
                            //    String hyperlink = matcher.group(1);
                            //    processHyperlink(wordDocument, currentBlock,
                            //            field.secondSubrange(parentRange),
                            //            currentTableLevel, hyperlink);
                            //    return;
                            //}
                        }
                        break;
                    }
            }

            logger.Log(POILogger.WARN, parentRange + " contains " + field
                    + " with unsupported type or format");
            ProcessCharacters(wordDocument, currentTableLevel,
                    field.SecondSubrange(parentRange), currentBlock);
        }

        protected abstract void ProcessFootnoteAutonumbered(
                HWPFDocument wordDocument, int noteIndex, XmlElement block,
                Range footnoteTextRange);

        protected abstract void ProcessHyperlink(HWPFDocumentCore wordDocument,
                XmlElement currentBlock, Range textRange, int currentTableLevel,
                String hyperlink);

        protected abstract void ProcessImage(XmlElement currentBlock,
                bool inlined, Picture picture);

        protected abstract void ProcessLineBreak(XmlElement block,
                CharacterRun characterRun);

        protected void ProcessNoteAnchor(HWPFDocument doc, CharacterRun characterRun, XmlElement block)
        {
            {

                Notes footnotes = doc.GetFootnotes();
                int noteIndex = footnotes.GetNoteIndexByAnchorPosition(characterRun.StartOffset);
                if (noteIndex != -1)
                {
                    Range footnoteRange = doc.GetFootnoteRange();
                    int rangeStartOffset = footnoteRange.StartOffset;
                    int noteTextStartOffset = footnotes.GetNoteTextStartOffSet(noteIndex);
                    int noteTextEndOffset = footnotes.GetNoteTextEndOffSet(noteIndex);

                    Range noteTextRange = new Range(rangeStartOffset + noteTextStartOffset, rangeStartOffset + noteTextEndOffset, doc);

                    ProcessFootnoteAutonumbered(doc, noteIndex, block, noteTextRange);
                    return;
                }
            }
            {
                Notes endnotes = doc.GetEndnotes();
                int noteIndex = endnotes.GetNoteIndexByAnchorPosition(characterRun.StartOffset);
                if (noteIndex != -1)
                {
                    Range endnoteRange = doc.GetEndnoteRange();
                    int rangeStartOffset = endnoteRange.StartOffset;
                    int noteTextStartOffset = endnotes.GetNoteTextStartOffSet(noteIndex);
                    int noteTextEndOffset = endnotes.GetNoteTextEndOffSet(noteIndex);

                    Range noteTextRange = new Range(rangeStartOffset + noteTextStartOffset, rangeStartOffset + noteTextEndOffset, doc);

                    ProcessEndnoteAutonumbered(doc, noteIndex, block, noteTextRange);
                    return;
                }
            }
            throw new NotImplementedException();
        }

        private bool ProcessOle2(HWPFDocument doc, CharacterRun characterRun,
                XmlElement block)
        {
            Entry entry = doc.GetObjectsPool().GetObjectById("_" + characterRun.GetPicOffset());
            if (entry == null)
            {
                logger.Log(POILogger.WARN, "Referenced OLE2 object '", (characterRun.GetPicOffset()).ToString(), "' not found in ObjectPool");
                return false;
            }

            try
            {
                return ProcessOle2(doc, block, entry);
            }
            catch (Exception exc)
            {
                logger.Log(POILogger.WARN,
                        "Unable to convert internal OLE2 object '", (characterRun.GetPicOffset()).ToString(), "': ", exc, exc);
                return false;
            }
        }

        ////@SuppressWarnings( "unused" )
        protected virtual bool ProcessOle2(HWPFDocument wordDocument, XmlElement block,
                NPOI.POIFS.FileSystem.Entry entry)
        {
            return false;
        }

        protected abstract void ProcessPageBreak(HWPFDocumentCore wordDocument,
                XmlElement flow);

        protected abstract void ProcessPageref(HWPFDocumentCore wordDocument,
                XmlElement currentBlock, Range textRange, int currentTableLevel,
                String pageref);

        protected abstract void ProcessParagraph(HWPFDocumentCore wordDocument,
                XmlElement parentElement, int currentTableLevel, Paragraph paragraph,
                String bulletText);

        protected void ProcessParagraphes(HWPFDocumentCore wordDocument,
                XmlElement flow, Range range, int currentTableLevel)
        {

            ListTables listTables = wordDocument.GetListTables();
            int currentListInfo = 0;

            int paragraphs = range.NumParagraphs;
            for (int p = 0; p < paragraphs; p++)
            {
                Paragraph paragraph = range.GetParagraph(p);

                if (paragraph.IsInTable() && paragraph.GetTableLevel() != currentTableLevel)
                {
                    if (paragraph.GetTableLevel() < currentTableLevel)
                        throw new InvalidOperationException(
                                "Trying to process table cell with higher level ("
                                        + paragraph.GetTableLevel()
                                        + ") than current table level ("
                                        + currentTableLevel
                                        + ") as inner table part");

                    Table table = range.GetTable(paragraph);
                    ProcessTable(wordDocument, flow, table);

                    p += table.NumParagraphs;
                    p--;
                    continue;
                }

                if (paragraph.Text.Equals("\u000c"))
                {
                    ProcessPageBreak(wordDocument, flow);
                }

                if (paragraph.GetIlfo() != currentListInfo)
                {
                    currentListInfo = paragraph.GetIlfo();
                }

                if (currentListInfo != 0)
                {
                    if (listTables != null)
                    {
                        ListFormatOverride listFormatOverride = listTables.GetOverride(paragraph.GetIlfo());

                        String label = AbstractWordUtils.GetBulletText(listTables,
                                paragraph, listFormatOverride.GetLsid());

                        ProcessParagraph(wordDocument, flow, currentTableLevel,
                                paragraph, label);
                    }
                    else
                    {
                        logger.Log(POILogger.WARN,
                                "Paragraph #" + paragraph.StartOffset + "-"
                                        + paragraph.EndOffset
                                        + " has reference to list structure #"
                                        + currentListInfo
                                        + ", but listTables not defined in file");

                        ProcessParagraph(wordDocument, flow, currentTableLevel,
                                paragraph, string.Empty);
                    }
                }
                else
                {
                    ProcessParagraph(wordDocument, flow, currentTableLevel,
                            paragraph, string.Empty);
                }
            }
        }

        protected abstract void ProcessSection(HWPFDocumentCore wordDocument,
                Section section, int s);

        protected virtual void ProcessSingleSection(HWPFDocumentCore wordDocument,
                Section section)
        {
            ProcessSection(wordDocument, section, 0);
        }

        protected abstract void ProcessTable(HWPFDocumentCore wordDocument,
                XmlElement flow, Table table);

        public void SetFontReplacer(FontReplacer fontReplacer)
        {
            this.fontReplacer = fontReplacer;
        }

        public void SetPicturesManager(PicturesManager fileManager)
        {
            this.picturesManager = fileManager;
        }

        protected int TryDeadField(HWPFDocumentCore wordDocument, Range range,
                int currentTableLevel, int beginMark, XmlElement currentBlock)
        {
            int separatorMark = -1;
            int endMark = -1;
            for (int c = beginMark + 1; c < range.NumCharacterRuns; c++)
            {
                CharacterRun characterRun = range.GetCharacterRun(c);

                String text = characterRun.Text;
                byte[] textByte = System.Text.Encoding.GetEncoding("iso-8859-1").GetBytes(text);
                if (textByte.Length == 0)
                    //if (text.getBytes().length == 0)
                    continue;

                if (textByte[0] == FIELD_BEGIN_MARK)
                {
                    // nested?
                    Field possibleField = ProcessDeadField(wordDocument, range,
                            currentTableLevel, characterRun.StartOffset,
                            currentBlock);
                    if (possibleField != null)
                    {
                        c = possibleField.GetFieldEndOffset();
                    }
                    else
                    {
                        continue;
                    }
                }

                if (textByte[0] == FIELD_SEPARATOR_MARK)
                {
                    if (separatorMark != -1)
                    {
                        // double;
                        return beginMark;
                    }

                    separatorMark = c;
                    continue;
                }

                if (textByte[0] == FIELD_END_MARK)
                {
                    if (endMark != -1)
                    {
                        // double;
                        return beginMark;
                    }

                    endMark = c;
                    break;
                }

            }

            if (separatorMark == -1 || endMark == -1)
                return beginMark;

            ProcessDeadField(wordDocument, currentBlock, range, currentTableLevel,
                    beginMark, separatorMark, endMark);

            return endMark;
        }
    }

}
