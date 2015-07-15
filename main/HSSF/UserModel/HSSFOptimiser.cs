/* ====================================================================
   Copyright 2002-2004   Apache Software Foundation

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
namespace NPOI.HSSF.UserModel
{
    using System.Collections;
    using NPOI.SS.UserModel;

    using NPOI.HSSF.Record;

    /// <summary>
    /// Excel can Get cranky if you give it files containing too
    /// many (especially duplicate) objects, and this class can
    /// help to avoid those.
    /// In general, it's much better to make sure you don't
    /// duplicate the objects in your code, as this is likely
    /// to be much faster than creating lots and lots of
    /// excel objects+records, only to optimise them down to
    /// many fewer at a later stage.
    /// However, sometimes this is too hard / tricky to do, which
    /// is where the use of this class comes in.
    /// </summary>
    public class HSSFOptimiser
    {
        /// <summary>
        /// Goes through the Workbook, optimising the fonts by
        /// removing duplicate ones.
        /// For now, only works on fonts used in HSSFCellStyle
        /// and HSSFRichTextString. Any other font uses
        /// (eg charts, pictures) may well end up broken!
        /// This can be a slow operation, especially if you have
        /// lots of cells, cell styles or rich text strings
        /// </summary>
        /// <param name="workbook">The workbook in which to optimise the fonts</param>
        public static void OptimiseFonts(HSSFWorkbook workbook)
        {
            // Where each font has ended up, and if we need to
            //  delete the record for it. Start off with no change
            short[] newPos =
                new short[workbook.Workbook.NumberOfFontRecords + 1];
            bool[] zapRecords = new bool[newPos.Length];
            for (int i = 0; i < newPos.Length; i++)
            {
                newPos[i] = (short)i;
                zapRecords[i] = false;
            }

            // Get each font record, so we can do deletes
            //  without Getting confused
            FontRecord[] frecs = new FontRecord[newPos.Length];
            for (int i = 0; i < newPos.Length; i++)
            {
                // There is no 4!
                if (i == 4) continue;

                frecs[i] = workbook.Workbook.GetFontRecordAt(i);
            }

            // Loop over each font, seeing if it is the same
            //  as an earlier one. If it is, point users of the
            //  later duplicate copy to the earlier one, and 
            //  mark the later one as needing deleting
            // Note - don't change built in fonts (those before 5)
            for (int i = 5; i < newPos.Length; i++)
            {
                // Check this one for being a duplicate
                //  of an earlier one
                int earlierDuplicate = -1;
                for (int j = 0; j < i && earlierDuplicate == -1; j++)
                {
                    if (j == 4) continue;

                    FontRecord frCheck = workbook.Workbook.GetFontRecordAt(j);
                    if (frCheck.SameProperties(frecs[i]))
                    {
                        earlierDuplicate = j;
                    }
                }

                // If we got a duplicate, mark it as such
                if (earlierDuplicate != -1)
                {
                    newPos[i] = (short)earlierDuplicate;
                    zapRecords[i] = true;
                }
            }

            // Update the new positions based on
            //  deletes that have occurred between
            //  the start and them
            // Only need to worry about user fonts
            for (int i = 5; i < newPos.Length; i++)
            {
                // Find the number deleted to that
                //  point, and adjust
                short preDeletePos = newPos[i];
                short newPosition = preDeletePos;
                for (int j = 0; j < preDeletePos; j++)
                {
                    if (zapRecords[j]) newPosition--;
                }

                // Update the new position
                newPos[i] = newPosition;
            }

            // Zap the un-needed user font records
            for (int i = 5; i < newPos.Length; i++)
            {
                if (zapRecords[i])
                {
                    workbook.Workbook.RemoveFontRecord(
                            frecs[i]
                    );
                }
            }

            // Tell HSSFWorkbook that it needs to
            //  re-start its HSSFFontCache
            workbook.ResetFontCache();

            // Update the cell styles to point at the 
            //  new locations of the fonts
            for (int i = 0; i < workbook.Workbook.NumExFormats; i++)
            {
                ExtendedFormatRecord xfr = workbook.Workbook.GetExFormatAt(i);
                xfr.FontIndex = (
                        newPos[xfr.FontIndex]
                );
            }

            // Update the rich text strings to point at
            //  the new locations of the fonts
            // Remember that one underlying unicode string
            //  may be shared by multiple RichTextStrings!
            ArrayList doneUnicodeStrings = new ArrayList();
            for (int sheetNum = 0; sheetNum < workbook.NumberOfSheets; sheetNum++)
            {
                NPOI.SS.UserModel.ISheet s = workbook.GetSheetAt(sheetNum);
                //IEnumerator rIt = s.GetRowEnumerator();
                //while (rIt.MoveNext())
                foreach (IRow row in s) 
                {
                    //HSSFRow row = (HSSFRow)rIt.Current;
                    //IEnumerator cIt = row.GetEnumerator();
                    //while (cIt.MoveNext())
                    foreach (ICell cell in row) 
                    {
                        //ICell cell = (HSSFCell)cIt.Current;
                        if (cell.CellType == NPOI.SS.UserModel.CellType.String)
                        {
                            HSSFRichTextString rtr = (HSSFRichTextString)cell.RichStringCellValue;
                            UnicodeString u = rtr.RawUnicodeString;

                            // Have we done this string already?
                            if (!doneUnicodeStrings.Contains(u))
                            {
                                // Update for each new position
                                for (short i = 5; i < newPos.Length; i++)
                                {
                                    if (i != newPos[i])
                                    {
                                        u.SwapFontUse(i, newPos[i]);
                                    }
                                }

                                // Mark as done
                                doneUnicodeStrings.Add(u);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Goes through the Wokrbook, optimising the cell styles
        /// by removing duplicate ones and ones that aren't used.
        /// For best results, optimise the fonts via a call to
        /// OptimiseFonts(HSSFWorkbook) first
        /// </summary>
        /// <param name="workbook">The workbook in which to optimise the cell styles</param>
        public static void OptimiseCellStyles(HSSFWorkbook workbook)
        {
            // Where each style has ended up, and if we need to
            //  delete the record for it. Start off with no change
            short[] newPos =
                new short[workbook.Workbook.NumExFormats];
            bool[] isUsed = new bool[newPos.Length];
            bool[] zapRecords = new bool[newPos.Length];
            for (int i = 0; i < newPos.Length; i++)
            {
                isUsed[i] = false;
                newPos[i] = (short)i;
                zapRecords[i] = false;
            }

            // Get each style record, so we can do deletes
            //  without Getting confused
            ExtendedFormatRecord[] xfrs = new ExtendedFormatRecord[newPos.Length];
            for (int i = 0; i < newPos.Length; i++)
            {
                xfrs[i] = workbook.Workbook.GetExFormatAt(i);
            }

            // Loop over each style, seeing if it is the same
            //  as an earlier one. If it is, point users of the
            //  later duplicate copy to the earlier one, and 
            //  mark the later one as needing deleting
            // Only work on user added ones, which come after 20
            for (int i = 21; i < newPos.Length; i++)
            {
                // Check this one for being a duplicate
                //  of an earlier one
                int earlierDuplicate = -1;
                for (int j = 0; j < i && earlierDuplicate == -1; j++)
                {
                    ExtendedFormatRecord xfCheck = workbook.Workbook.GetExFormatAt(j);
                    if (xfCheck.Equals(xfrs[i]))
                    {
                        earlierDuplicate = j;
                    }
                }

                // If we got a duplicate, mark it as such
                if (earlierDuplicate != -1)
                {
                    newPos[i] = (short)earlierDuplicate;
                    zapRecords[i] = true;
                }
                // If we got a duplicate, mark the one we're keeping as used
                if (earlierDuplicate != -1)
                {
                    isUsed[earlierDuplicate] = true;
                }
            }
            // Loop over all the cells in the file, and identify any user defined
            //  styles aren't actually being used (don't touch built-in ones)
            for (int sheetNum = 0; sheetNum < workbook.NumberOfSheets; sheetNum++)
            {
                HSSFSheet s = (HSSFSheet)workbook.GetSheetAt(sheetNum);
                foreach (IRow row in s)
                {
                    foreach (ICell cellI in row)
                    {
                        HSSFCell cell = (HSSFCell)cellI;
                        short oldXf = cell.CellValueRecord.XFIndex;
                        isUsed[oldXf] = true;
                    }
                }
            }
            // Mark any that aren't used as needing zapping
            for (int i = 21; i < isUsed.Length; i++)
            {
                if (!isUsed[i])
                {
                    // Un-used style, can be removed
                    zapRecords[i] = true;
                    newPos[i] = 0;
                }
            }
            // Update the new positions based on
            //  deletes that have occurred between
            //  the start and them
            // Only work on user added ones, which come after 20
            for (int i = 21; i < newPos.Length; i++)
            {
                // Find the number deleted to that
                //  point, and adjust
                short preDeletePos = newPos[i];
                short newPosition = preDeletePos;
                for (int j = 0; j < preDeletePos; j++)
                {
                    if (zapRecords[j]) newPosition--;
                }

                // Update the new position
                newPos[i] = newPosition;
            }

            // Zap the un-needed user style records
            // removing by index, because removing by object may delete
            // styles we did not intend to (the ones that _were_ duplicated and not the duplicates)
            int max = newPos.Length;
            int removed = 0; // to adjust index after deletion
            for (int i = 21; i < max; i++)
            {
                if (zapRecords[i + removed])
                {
                    workbook.Workbook.RemoveExFormatRecord(i);
                    i--;
                    max--;
                    removed++;
                }
            }

            // Finally, update the cells to point at their new extended format records
            for (int sheetNum = 0; sheetNum < workbook.NumberOfSheets; sheetNum++)
            {
                HSSFSheet s = (HSSFSheet)workbook.GetSheetAt(sheetNum);
                //IEnumerator rIt = s.GetRowEnumerator();
                //while (rIt.MoveNext())
                foreach(IRow row in s)
                {
                    //HSSFRow row = (HSSFRow)rIt.Current;
                    //IEnumerator cIt = row.GetEnumerator();
                    //while (cIt.MoveNext())
                    foreach (ICell cell in row)
                    {
                        //ICell cell = (HSSFCell)cIt.Current;
                        short oldXf = ((HSSFCell)cell).CellValueRecord.XFIndex;
                        NPOI.SS.UserModel.ICellStyle newStyle = workbook.GetCellStyleAt(
                                newPos[oldXf]
                        );
                        cell.CellStyle = (newStyle);
                    }
                }
            }
        }
    }
}