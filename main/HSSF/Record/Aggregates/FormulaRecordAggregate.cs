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
    using NPOI.HSSF.Record;
    using NPOI.SS.Util;
    using NPOI.SS.Formula.PTG;
    using NPOI.Util;

    /// <summary>
    /// The formula record aggregate is used to join toGether the formula record and it's
    /// (optional) string record and (optional) Shared Formula Record (template Reads, excel optimization).
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    [Serializable]
    public class FormulaRecordAggregate: RecordAggregate, CellValueRecordInterface, IComparable, ICloneable
    {
        public const short sid = -2000;

        private FormulaRecord _formulaRecord;
        private SharedValueManager _sharedValueManager;
        /** caches the calculated result of the formula */
        private StringRecord _stringRecord;
        [NonSerialized]
        private SharedFormulaRecord _sharedFormulaRecord;
        /// <summary>
        /// Initializes a new instance of the <see cref="FormulaRecordAggregate"/> class.
        /// </summary>
        /// <param name="formulaRec">The formula rec.</param>
        /// <param name="stringRec">The string rec.</param>
        /// <param name="svm">The SVM.</param>
        public FormulaRecordAggregate(FormulaRecord formulaRec, StringRecord stringRec, SharedValueManager svm)
        {
            if (svm == null)
            {
                throw new ArgumentException("sfm must not be null");
            }
            if (formulaRec.HasCachedResultString)
            {
                if (stringRec == null)
                {
                    throw new RecordFormatException("Formula record flag is set but String record was not found");
                }
                _stringRecord = stringRec;
            }
            else
            {
                // Usually stringRec is null here (in agreement with what the formula rec says).
                // In the case where an extra StringRecord is erroneously present, Excel (2007)
                // ignores it (see bug 46213).
                _stringRecord = null;
            }

            _formulaRecord = formulaRec;
            _sharedValueManager = svm;
            if (formulaRec.IsSharedFormula)
            {
                CellReference firstCell = formulaRec.Formula.ExpReference;
                if (firstCell == null)
                {
                    HandleMissingSharedFormulaRecord(formulaRec);
                }
                else
                {
                    _sharedFormulaRecord = svm.LinkSharedFormulaRecord(firstCell, this);
                }
            }
        }
        	/**
	     * Should be called by any code which is either deleting this formula cell, or changing
	     * its type.  This method gives the aggregate a chance to unlink any shared formula
	     * that may be involved with this cell formula.
	     */
	    public void NotifyFormulaChanging() {
		    if (_sharedFormulaRecord != null) {
			    _sharedValueManager.Unlink(_sharedFormulaRecord);
		    }
	    }
        public bool IsPartOfArrayFormula
        {
            get
            {
                if (_sharedFormulaRecord != null)
                {
                    return false;
                }
                CellReference expRef = _formulaRecord.Formula.ExpReference;
                ArrayRecord arec = expRef == null ? null : _sharedValueManager.GetArrayRecord(expRef.Row, expRef.Col);
                return arec != null;
            }
        }

        /// <summary>
        /// called by the class that is responsible for writing this sucker.
        /// Subclasses should implement this so that their data is passed back in a
        /// byte array.
        /// </summary>
        /// <param name="offset">offset to begin writing at</param>
        /// <param name="data">byte array containing instance data.</param>
        /// <returns>number of bytes written</returns>
        public override int Serialize(int offset, byte [] data)
        {
            int pos = offset;
            pos += _formulaRecord.Serialize(pos, data);

            if (_stringRecord != null)
            {
                pos += _stringRecord.Serialize(pos, data);
            }
            return pos - offset;

        }
        /// <summary>
        /// Visit each of the atomic BIFF records contained in this {@link RecordAggregate} in the order
        /// that they should be written to file.  Implementors may or may not return the actual
        /// {@link Record}s being used to manage POI's internal implementation.  Callers should not
        /// assume either way, and therefore only attempt to modify those {@link Record}s after cloning
        /// </summary>
        /// <param name="rv"></param>
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(_formulaRecord);
            Record sharedFormulaRecord = _sharedValueManager.GetRecordForFirstCell(this);
            if (sharedFormulaRecord != null)
            {
                rv.VisitRecord(sharedFormulaRecord);
            }
            if (_formulaRecord.HasCachedResultString && _stringRecord != null)
            {
                rv.VisitRecord(_stringRecord);
            }
        }
        /// <summary>
        /// Get the current Serialized size of the record. Should include the sid and recLength (4 bytes).
        /// </summary>
        /// <value>The size of the record.</value>
        public override int RecordSize
        {
            get
            {
                int size = _formulaRecord.RecordSize + (_stringRecord == null ? 0 : _stringRecord.RecordSize);
                return size;
            }
        }


        /// <summary>
        /// return the non static version of the id for this record.
        /// </summary>
        /// <value>The sid.</value>
        public override short Sid
        {
            get
            {
                return sid;
            }
        }
        /// <summary>
        /// Sometimes the shared formula flag "seems" to be erroneously set (because the corresponding
        /// SharedFormulaRecord does not exist). Normally this would leave no way of determining
        /// the Ptg tokens for the formula.  However as it turns out in these
        /// cases, Excel encodes the unshared Ptg tokens in the right place (inside the FormulaRecord). 
        /// So the the only thing that needs to be done is to ignore the erroneous
        /// shared formula flag.
        /// 
        /// This method may also be used for setting breakpoints to help diagnose issues regarding the
        /// abnormally-set 'shared formula' flags.
        /// </summary>
        /// <param name="formula">The formula.</param>
        private static void HandleMissingSharedFormulaRecord(FormulaRecord formula)
        {
            // make sure 'unshared' formula is actually available
            Ptg firstToken = formula.ParsedExpression[0];
            if (firstToken is ExpPtg)
            {
                throw new RecordFormatException(
                        "SharedFormulaRecord not found for FormulaRecord with (isSharedFormula=true)");
            }
            // could log an info message here since this is a fairly unusual occurrence.
            formula.IsSharedFormula = false; // no point leaving the flag erroneously set
        }

        /// <summary>
        /// Gets or sets the formula record.
        /// </summary>
        /// <value>The formula record.</value>
        public FormulaRecord FormulaRecord
        {
            get
            {
                return _formulaRecord;
            }
            set { this._formulaRecord = value; }
        }

        /// <summary>
        /// Gets or sets the string record.
        /// </summary>
        /// <value>The string record.</value>
        public StringRecord StringRecord
        {
            get
            {
                return _stringRecord;
            }
            set { this._stringRecord = value; }
        }


        public short XFIndex
        {
            get{return _formulaRecord.XFIndex;}
            set{_formulaRecord.XFIndex=value;}
        }

        public int Column
        {
            get{return _formulaRecord.Column;}
            set{_formulaRecord.Column=value;}
        }

        public int Row
        {
            get { return _formulaRecord.Row; }
            set { _formulaRecord.Row=value; }
        }



        public int CompareTo(Object o)
        {
            return _formulaRecord.CompareTo(o);
        }

        ///// <summary>
        ///// returns whether this cell represents the same cell (NOT VALUE)
        ///// </summary>
        ///// <param name="i"> record to Compare</param>
        ///// <returns>true if the cells are the same cell (positionally), false if not.</returns>
        //public bool IsEqual(CellValueRecordInterface i)
        //{
        //    return _formulaRecord.IsEqual(i);
        //}


        /// <summary>
        /// Determines whether the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <param name="obj">The <see cref="T:System.Object"/> to compare with the current <see cref="T:System.Object"/>.</param>
        /// <returns>
        /// true if the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>; otherwise, false.
        /// </returns>
        /// <exception cref="T:System.NullReferenceException">
        /// The <paramref name="obj"/> parameter is null.
        /// </exception>
        public override bool Equals(Object obj)
        {
            return _formulaRecord.Equals(obj);
        }

        public override int GetHashCode ()
        {
            return _formulaRecord.GetHashCode ();
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            return _formulaRecord.ToString();
        }

        /// <summary>
        /// Gets the string value.
        /// </summary>
        /// <value>The string value.</value>
        public String StringValue
        {
            get
            {
                if (_stringRecord == null) return null;
                return _stringRecord.String;
            }
        }
        public void SetCachedDoubleResult(double value)
        {
            _stringRecord = null;
            _formulaRecord.Value = value;
        }

        /// <summary>
        /// Sets the cached string result.
        /// </summary>
        /// <param name="value">The value.</param>
        public void SetCachedStringResult(String value)
        {

            // Save the string into a String Record, creating one if required
            if (_stringRecord == null)
            {
                _stringRecord = new StringRecord();
            }
            _stringRecord.String=(value);
            if (value.Length < 1)
            {
                _formulaRecord.SetCachedResultTypeEmptyString();
            }
            else
            {
                _formulaRecord.SetCachedResultTypeString();
            }
        }
        /// <summary>
        /// Sets the cached boolean result.
        /// </summary>
        /// <param name="value">if set to <c>true</c> [value].</param>
        public void SetCachedBooleanResult(bool value)
        {
            _stringRecord = null;
            _formulaRecord.SetCachedResultBoolean(value);
        }
        /// <summary>
        /// Sets the cached error result.
        /// </summary>
        /// <param name="errorCode">The error code.</param>
        public void SetCachedErrorResult(int errorCode)
        {
            _stringRecord = null;
            _formulaRecord.SetCachedResultErrorCode(errorCode);
        }

        #region ICloneable Members

        public object Clone()
        {
            return this;
        }

        #endregion
        public Ptg[] FormulaTokens
        {
            get
            {
                if (_sharedFormulaRecord != null)
                {
                    return _sharedFormulaRecord.GetFormulaTokens(_formulaRecord);
                }
                CellReference expRef = _formulaRecord.Formula.ExpReference;
                if (expRef != null)
                {
                    ArrayRecord arec = _sharedValueManager.GetArrayRecord(expRef.Row, expRef.Col);
                    return arec.FormulaTokens;
                }
                return _formulaRecord.ParsedExpression;
            }
        }
        /**
 * Also checks for a related shared formula and unlinks it if found
 */
        public void SetParsedExpression(Ptg[] ptgs)
        {
            NotifyFormulaChanging();
            _formulaRecord.ParsedExpression=(ptgs);
        }
        public void UnlinkSharedFormula()
        {
            SharedFormulaRecord sfr = _sharedFormulaRecord;
            if (sfr == null)
            {
                throw new InvalidOperationException("Formula not linked to shared formula");
            }
            Ptg[] ptgs = sfr.GetFormulaTokens(_formulaRecord);
            _formulaRecord.SetParsedExpression(ptgs);
            //Now its not shared!
            _formulaRecord.SetSharedFormula(false);
            _sharedFormulaRecord = null;
        }
        public CellRangeAddress GetArrayFormulaRange()
        {
            if (_sharedFormulaRecord != null)
            {
                throw new InvalidOperationException("not an array formula cell.");
            }
            CellReference expRef = _formulaRecord.Formula.ExpReference;
            if (expRef == null)
            {
                throw new InvalidOperationException("not an array formula cell.");
            }
            ArrayRecord arec = _sharedValueManager.GetArrayRecord(expRef.Row, expRef.Col);
            if (arec == null)
            {
                throw new InvalidOperationException("ArrayRecord was not found for the locator " + expRef.FormatAsString());
            }
            CellRangeAddress8Bit a = arec.Range;
            return new CellRangeAddress(a.FirstRow, a.LastRow, a.FirstColumn, a.LastColumn);
        }
        public void SetArrayFormula(CellRangeAddress r, Ptg[] ptgs)
        {

            ArrayRecord arr = new ArrayRecord(NPOI.SS.Formula.Formula.Create(ptgs), new CellRangeAddress8Bit(r.FirstRow, r.LastRow, r.FirstColumn, r.LastColumn));
            _sharedValueManager.AddArrayRecord(arr);
        }
        /**
         * Removes an array formula
         * @return the range of the array formula containing the specified cell. Never <code>null</code>
         */
        public CellRangeAddress RemoveArrayFormula(int rowIndex, int columnIndex)
        {
            CellRangeAddress8Bit a = _sharedValueManager.RemoveArrayFormula(rowIndex, columnIndex);
            // at this point FormulaRecordAggregate#isPartOfArrayFormula() should return false
            _formulaRecord.ParsedExpression = (null);
            return new CellRangeAddress(a.FirstRow, a.LastRow, a.FirstColumn, a.LastColumn);
        }
    }
}