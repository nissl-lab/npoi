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

namespace NPOI.HSSF.Record.Aggregates
{
    using System;
    using System.Text;
    using System.Collections;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Model;
    using NPOI.SS.Formula;
    using System.Collections.Generic;
    using NPOI.SS.Util;
    using NPOI.SS.Formula.PTG;
    using NPOI.Util;

    /// <summary>
    /// 
    /// </summary>
    /// CFRecordsAggregate - aggregates Conditional Formatting records CFHeaderRecord
    /// and number of up to three CFRuleRecord records toGether to simplify
    /// access to them.
    /// @author Dmitriy Kumshayev
    public class CFRecordsAggregate : RecordAggregate
    {
        /** Excel allows up to 3 conditional formating rules */
        private const int MAX_97_2003_CONDTIONAL_FORMAT_RULES = 3;

        public const short sid = -2008; // not a real BIFF record

        //private static POILogger log = POILogFactory.GetLogger(typeof(CFRecordsAggregate));

        private CFHeaderBase header;

        /** List of CFRuleRecord objects */
        private List<CFRuleBase> rules;

        private CFRecordsAggregate(CFHeaderBase pHeader, CFRuleBase[] pRules)
        {
            if (pHeader == null)
            {
                throw new ArgumentException("header must not be null");
            }
            if (pRules == null)
            {
                throw new ArgumentException("rules must not be null");
            }
            if (pRules.Length > MAX_97_2003_CONDTIONAL_FORMAT_RULES)
            {
                Console.WriteLine("Excel versions before 2007 require that "
                    + "No more than " + MAX_97_2003_CONDTIONAL_FORMAT_RULES
                    + " rules may be specified, " + pRules.Length + " were found,"
                    + " this file will cause problems with old Excel versions");
            }
            if (pRules.Length != pHeader.NumberOfConditionalFormats)
            {
                throw new RecordFormatException("Mismatch number of rules");
            }
            header = pHeader;
            rules = new List<CFRuleBase>(pRules.Length);
            foreach (CFRuleBase pRule in pRules)
            {
                CheckRuleType(pRule);
                rules.Add(pRule);
            }
        }


        public CFRecordsAggregate(CellRangeAddress[] regions, CFRuleBase[] rules)
            : this(CreateHeader(regions, rules), rules)
        {

        }
        private static CFHeaderBase CreateHeader(CellRangeAddress[] regions, CFRuleBase[] rules)
        {
            CFHeaderBase header;
            if (rules.Length == 0 || rules[0] is CFRuleRecord) {
                header = new CFHeaderRecord(regions, rules.Length);
            }
            else
            {
                header = new CFHeader12Record(regions, rules.Length);
            }
            // set the "needs recalculate" by default to avoid Excel handling conditional formatting incorrectly
            // see bug 52122 for details
            header.NeedRecalculation = true;

            return header;
        }
        /// <summary>
        /// Create CFRecordsAggregate from a list of CF Records
        /// </summary>
        /// <param name="rs">list of Record objects</param>
        public static CFRecordsAggregate CreateCFAggregate(RecordStream rs)
        {
            Record rec = rs.GetNext();
            if (rec.Sid != CFHeaderRecord.sid &&
                rec.Sid != CFHeader12Record.sid)
            {
                throw new InvalidOperationException("next record sid was " + rec.Sid
                        + " instead of " + CFHeaderRecord.sid + " or " 
                        + CFHeader12Record.sid + " as expected");
            }

            CFHeaderBase header = (CFHeaderBase)rec;
            int nRules = header.NumberOfConditionalFormats;

            CFRuleBase[] rules = new CFRuleBase[nRules];
            for (int i = 0; i < rules.Length; i++)
            {
                rules[i] = (CFRuleBase)rs.GetNext();
            }

            return new CFRecordsAggregate(header, rules);
        }
        /// <summary>
        /// Create CFRecordsAggregate from a list of CF Records
        /// </summary>
        /// <param name="recs">list of Record objects</param>
        /// <param name="pOffset">position of CFHeaderRecord object in the list of Record objects</param>
        [Obsolete("Not found in poi(2015-07-14), maybe was removed")]
        public static CFRecordsAggregate CreateCFAggregate(IList recs, int pOffset)
        {
            Record rec = (Record)recs[pOffset];
            if (rec.Sid != CFHeaderRecord.sid)
            {
                throw new InvalidOperationException("next record sid was " + rec.Sid
                        + " instead of " + CFHeaderRecord.sid + " as expected");
            }

            CFHeaderRecord header = (CFHeaderRecord)rec;
            int nRules = header.NumberOfConditionalFormats;

            CFRuleRecord[] rules = new CFRuleRecord[nRules];
            int offset = pOffset;
            int countFound = 0;
            while (countFound < rules.Length)
            {
                offset++;
                if (offset >= recs.Count)
                {
                    break;
                }
                rec = (Record)recs[offset];
                if (rec is CFRuleRecord record)
                {
                    rules[countFound] = record;
                    countFound++;
                }
                else
                {
                    break;
                }
            }

            if (countFound < nRules)
            { // TODO -(MAR-2008) can this ever happen? Write junit 

                //if (log.Check(POILogger.DEBUG))
                //{
                //    log.Log(POILogger.DEBUG, "Expected  " + nRules + " Conditional Formats, "
                //            + "but found " + countFound + " rules");
                //}
                header.NumberOfConditionalFormats = (nRules);
                CFRuleRecord[] lessRules = new CFRuleRecord[countFound];
                Array.Copy(rules, 0, lessRules, 0, countFound);
                rules = lessRules;
            }
            return new CFRecordsAggregate(header, rules);
        }
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(header);
            foreach (CFRuleBase rule in rules)
            {
                rv.VisitRecord(rule);
            }
        }

        /// <summary>
        /// Create a deep Clone of the record
        /// </summary>
        public CFRecordsAggregate CloneCFAggregate()
        {

            CFRuleBase[] newRecs = new CFRuleBase[rules.Count];
            for (int i = 0; i < newRecs.Length; i++)
            {
                newRecs[i] = (CFRuleRecord)GetRule(i).Clone();
            }
            return new CFRecordsAggregate((CFHeaderBase)header.Clone(), newRecs);
        }

        public override short Sid
        {
            get { return sid; }
        }

        /// <summary>
        /// called by the class that is responsible for writing this sucker.
        /// Subclasses should implement this so that their data is passed back in a
        /// byte array.
        /// </summary>
        /// <param name="offset">The offset to begin writing at</param>
        /// <param name="data">The data byte array containing instance data</param>
        /// <returns> number of bytes written</returns>
        public override int Serialize(int offset, byte[] data)
        {
            int nRules = rules.Count;
            header.NumberOfConditionalFormats = (nRules);

            int pos = offset;

            pos += header.Serialize(pos, data);
            for (int i = 0; i < nRules; i++)
            {
                pos += GetRule(i).Serialize(pos, data);
            }
            return pos - offset;
        }

        public CFHeaderBase Header
        {
            get { return header; }
        }

        private void CheckRuleIndex(int idx)
        {
            if (idx < 0 || idx >= rules.Count)
            {
                throw new ArgumentException("Bad rule record index (" + idx
                        + ") nRules=" + rules.Count);
            }
        }
        private void CheckRuleType(CFRuleBase r)
        {
            if (header is CFHeaderRecord &&
                     r is CFRuleRecord) {
                return;
            }
            if (header is CFHeader12Record &&
                     r is CFRule12Record) {
                return;
            }
            throw new ArgumentException("Header and Rule must both be CF or both be CF12, can't mix");
        }
        public CFRuleBase GetRule(int idx)
        {
            CheckRuleIndex(idx);
            return rules[idx];
        }
        public void SetRule(int idx, CFRuleBase r)
        {
            CheckRuleIndex(idx);
            CheckRuleType(r);
            rules[idx] = r;
        }

        /**
         * @return <c>false</c> if this whole {@link CFHeaderRecord} / {@link CFRuleRecord}s should be deleted
         */
        public bool UpdateFormulasAfterCellShift(FormulaShifter shifter, int currentExternSheetIx)
        {
            CellRangeAddress[] cellRanges = header.CellRanges;
            bool changed = false;
            List<CellRangeAddress> temp = new List<CellRangeAddress>();
            foreach (CellRangeAddress craOld in cellRanges)
            {
                CellRangeAddress craNew = ShiftRange(shifter, craOld, currentExternSheetIx);
                if (craNew == null)
                {
                    changed = true;
                    continue;
                }
                temp.Add(craNew);
                if (craNew != craOld)
                {
                    changed = true;
                }
            }

            if (changed)
            {
                int nRanges = temp.Count;
                if (nRanges == 0)
                {
                    return false;
                }
                CellRangeAddress[] newRanges = new CellRangeAddress[nRanges];
                newRanges = temp.ToArray();
                header.CellRanges = (newRanges);
            }

            foreach (CFRuleBase rule in rules)
            {
                Ptg[] ptgs;
                ptgs = rule.ParsedExpression1;
                if (ptgs != null && shifter.AdjustFormula(ptgs, currentExternSheetIx))
                {
                    rule.ParsedExpression1 = (ptgs);
                }
                ptgs = rule.ParsedExpression2;
                if (ptgs != null && shifter.AdjustFormula(ptgs, currentExternSheetIx))
                {
                    rule.ParsedExpression2 = (ptgs);
                }
                if (rule is CFRule12Record rule12)
                {
                    ptgs = rule12.ParsedExpressionScale;
                    if (ptgs != null && shifter.AdjustFormula(ptgs, currentExternSheetIx))
                    {
                        rule12.ParsedExpressionScale = (ptgs);
                    }
                }
        }
            return true;
        }
        private static CellRangeAddress ShiftRange(FormulaShifter shifter, CellRangeAddress cra, int currentExternSheetIx)
        {
            // FormulaShifter works well in terms of Ptgs - so convert CellRangeAddress to AreaPtg (and back) here
            AreaPtg aptg = new AreaPtg(cra.FirstRow, cra.LastRow, cra.FirstColumn, cra.LastColumn, false, false, false, false);
            Ptg[] ptgs = { aptg, };

            if (!shifter.AdjustFormula(ptgs, currentExternSheetIx))
            {
                return cra;
            }
            Ptg ptg0 = ptgs[0];
            if (ptg0 is AreaPtg bptg)
            {
                return new CellRangeAddress(bptg.FirstRow, bptg.LastRow, bptg.FirstColumn, bptg.LastColumn);
            }
            if (ptg0 is AreaErrPtg)
            {
                return null;
            }
            throw new InvalidCastException("Unexpected shifted ptg class (" + ptg0.GetType().Name + ")");
        }
        public void AddRule(CFRuleBase r)
        {
            if (rules.Count >= MAX_97_2003_CONDTIONAL_FORMAT_RULES)
            {
                Console.WriteLine("Excel versions before 2007 cannot cope with"
                    + " any more than " + MAX_97_2003_CONDTIONAL_FORMAT_RULES
                    + " - this file will cause problems with old Excel versions");
            }
            CheckRuleType(r);
            rules.Add(r);
            header.NumberOfConditionalFormats = (rules.Count);
        }
        public int NumberOfRules
        {
            get { return rules.Count; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            String type = "CF";
            if (header is CFHeader12Record) {
                type = "CF12";
            }
            buffer.Append("[").Append(type).Append("]\n");
            if (header != null)
            {
                buffer.Append(header.ToString());
            }
            foreach (CFRuleBase cfRule in rules)
            {
                if (cfRule != null)
                {
                    buffer.Append(cfRule.ToString());
                }
            }
            buffer.Append("[/CF]\n");
            return buffer.ToString();
        }
    }
}