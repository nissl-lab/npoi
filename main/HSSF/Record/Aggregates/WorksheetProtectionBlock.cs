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

namespace NPOI.HSSF.Record.Aggregates
{

    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    using System;

    /**
     * Groups the sheet protection records for a worksheet.
     * <p/>
     *
     * See OOO excelfileformat.pdf sec 4.18.2 'Sheet Protection in a Workbook
     * (BIFF5-BIFF8)'
     *
     * @author Josh Micich
     */
    public class WorksheetProtectionBlock : RecordAggregate
    {
        // Every one of these component records is optional
        // (The whole WorksheetProtectionBlock may not be present)
        private ProtectRecord _protectRecord;
        private ObjectProtectRecord _objectProtectRecord;
        private ScenarioProtectRecord _scenarioProtectRecord;
        private PasswordRecord _passwordRecord;

        /**
         * Creates an empty WorksheetProtectionBlock
         */
        public WorksheetProtectionBlock()
        {
            // all fields emptyv
        }

        /**
         * @return <c>true</c> if the specified Record sid is one belonging to
         *         the 'Page Settings Block'.
         */
        public static bool IsComponentRecord(int sid)
        {
            switch (sid)
            {
                case ProtectRecord.sid:
                case ObjectProtectRecord.sid:
                case ScenarioProtectRecord.sid:
                case PasswordRecord.sid:
                    return true;
            }
            return false;
        }

        private bool ReadARecord(RecordStream rs)
        {
            
            switch (rs.PeekNextSid())
            {
                case ProtectRecord.sid:
                    CheckNotPresent(_protectRecord);
                    _protectRecord = rs.GetNext() as ProtectRecord;
                    break;
                case ObjectProtectRecord.sid:
                    CheckNotPresent(_objectProtectRecord);
                    _objectProtectRecord = rs.GetNext() as ObjectProtectRecord;
                    break;
                case ScenarioProtectRecord.sid:
                    CheckNotPresent(_scenarioProtectRecord);
                    _scenarioProtectRecord = rs.GetNext() as ScenarioProtectRecord;
                    break;
                case PasswordRecord.sid:
                    CheckNotPresent(_passwordRecord);
                    _passwordRecord = rs.GetNext() as PasswordRecord;
                    break;
                default:
                    // all other record types are not part of the PageSettingsBlock
                    return false;
            }
            return true;
        }

        private void CheckNotPresent(Record rec)
        {
            if (rec != null)
            {
                throw new RecordFormatException("Duplicate WorksheetProtectionBlock record (sid=0x"
                        + StringUtil.ToHexString(rec.Sid) + ")");
            }
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            // Replicates record order from Excel 2007, though this is not critical

            VisitIfPresent(_protectRecord, rv);
            VisitIfPresent(_objectProtectRecord, rv);
            VisitIfPresent(_scenarioProtectRecord, rv);
            VisitIfPresent(_passwordRecord, rv);
        }

        private static void VisitIfPresent(Record r, RecordVisitor rv)
        {
            if (r != null)
            {
                rv.VisitRecord(r);
            }
        }

        public PasswordRecord GetPasswordRecord()
        {
            return _passwordRecord;
        }

        public ScenarioProtectRecord GetHCenter()
        {
            return _scenarioProtectRecord;
        }

        /**
         * This method Reads {@link WorksheetProtectionBlock} records from the supplied RecordStream
         * until the first non-WorksheetProtectionBlock record is encountered. As each record is Read,
         * it is incorporated into this WorksheetProtectionBlock.
         * <p/>
         * As per the OOO documentation, the protection block records can be expected to be written
         * toGether (with no intervening records), but earlier versions of POI (prior to Jun 2009)
         * didn't do this.  Workbooks with sheet protection Created by those earlier POI versions
         * seemed to be valid (Excel opens them OK). So PO allows continues to support Reading of files
         * with non continuous worksheet protection blocks.
         *
         * <p/>
         * <b>Note</b> - when POI Writes out this WorksheetProtectionBlock, the records will always be
         * written in one consolidated block (in the standard ordering) regardless of how scattered the
         * records were when they were originally Read.
         */
        public void AddRecords(RecordStream rs)
        {
            while (true)
            {
                if (!ReadARecord(rs))
                {
                    break;
                }
            }
        }

        /// <summary>
        /// the ProtectRecord. If one is not contained in the sheet, then one is created.
        /// </summary>
        private ProtectRecord Protect
        {
            get
            {
                if (_protectRecord == null)
                {
                    _protectRecord = new ProtectRecord(false);
                }
                return _protectRecord;
            }
        }
        /// <summary>
        /// the PasswordRecord. If one is not Contained in the sheet, then one is Created.
        /// </summary>
        public PasswordRecord Password
        {
            get
            {
                if (_passwordRecord == null)
                {
                    _passwordRecord = CreatePassword();
                }
                return _passwordRecord;
            }
        }

        /// <summary>
        /// protect a spreadsheet with a password (not encrypted, just sets protect flags and the password.)
        /// </summary>
        /// <param name="password">password to set;Pass <code>null</code> to remove all protection</param>
        /// <param name="shouldProtectObjects">shouldProtectObjects are protected</param>
        /// <param name="shouldProtectScenarios">shouldProtectScenarios are protected</param>
        public void ProtectSheet(String password, bool shouldProtectObjects,
                bool shouldProtectScenarios)
        {
            if (password == null)
            {
                _passwordRecord = null;
                _protectRecord = null;
                _objectProtectRecord = null;
                _scenarioProtectRecord = null;
                return;
            }

            ProtectRecord prec = this.Protect;
            PasswordRecord pass = this.Password;
            prec.Protect = true;
            pass.Password = (PasswordRecord.HashPassword(password));
            if (_objectProtectRecord == null && shouldProtectObjects)
            {
                ObjectProtectRecord rec = CreateObjectProtect();
                rec.Protect = (true);
                _objectProtectRecord = rec;
            }
            if (_scenarioProtectRecord == null && shouldProtectScenarios)
            {
                ScenarioProtectRecord srec = CreateScenarioProtect();
                srec.Protect = (true);
                _scenarioProtectRecord = srec;
            }
        }

        public bool IsSheetProtected
        {
            get
            {
                return _protectRecord != null && _protectRecord.Protect;
            }
        }

        public bool IsObjectProtected
        {
            get{
                return _objectProtectRecord != null && _objectProtectRecord.Protect;
            }
        }

        public bool IsScenarioProtected
        {
            get
            {
                return _scenarioProtectRecord != null && _scenarioProtectRecord.Protect;
            }
        }

        /// <summary>
        /// Creates an ObjectProtect record with protect set to false.
        /// </summary>
        /// <returns></returns>
        private static ObjectProtectRecord CreateObjectProtect() {
		ObjectProtectRecord retval = new ObjectProtectRecord();
		retval.Protect = (false);
		return retval;
	}
        /// <summary>
        /// Creates a ScenarioProtect record with protect set to false.
        /// </summary>
        /// <returns></returns>
        private static ScenarioProtectRecord CreateScenarioProtect() {
		ScenarioProtectRecord retval = new ScenarioProtectRecord();
		retval.Protect = (false);
		return retval;
	}

        /// <summary>
        ///Creates a Password record with password set to 0x0000.
        /// </summary>
        /// <returns></returns>
        private static PasswordRecord CreatePassword()
        {
            return new PasswordRecord(0x0000);
        }

        public int PasswordHash
        {
            get
            {
                if (_passwordRecord == null)
                {
                    return 0;
                }
                return _passwordRecord.Password;
            }
        }
    }
}



