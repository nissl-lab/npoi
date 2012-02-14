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

namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.Util;
    using NPOI.HSSF.Record;

    /**
     * Title:        Sup Book (EXTERNALBOOK) 
     * Description:  A External Workbook Description (Suplemental Book)
     *               Its only a dummy record for making new ExternSheet Record 
     * REFERENCE:  5.38
     * @author Libin Roman (Vista Portal LDT. Developer)
     * @author Andrew C. Oliver (acoliver@apache.org)
     *
     */
    public class SupBookRecord : StandardRecord
    {

        public const short sid = 0x1AE;

        private const short SMALL_RECORD_SIZE = 4;
        private const short TAG_INTERNAL_REFERENCES = 0x0401;
        private const short TAG_ADD_IN_FUNCTIONS = 0x3A01;

        private short field_1_number_of_sheets;
        private String field_2_encoded_url;
        private String[] field_3_sheet_names;
        private bool _isAddInFunctions;


        public static SupBookRecord CreateInternalReferences(short numberOfSheets)
        {
            return new SupBookRecord(false, numberOfSheets);
        }
        public static SupBookRecord CreateAddInFunctions()
        {
            return new SupBookRecord(true, (short)1);
        }
        public static SupBookRecord CreateExternalReferences(String url, String[] sheetNames)
        {
            return new SupBookRecord(url, sheetNames);
        }
        private SupBookRecord(bool IsAddInFuncs, short numberOfSheets)
        {
            // else not 'External References'
            field_1_number_of_sheets = numberOfSheets;
            field_2_encoded_url = null;
            field_3_sheet_names = null;
            _isAddInFunctions = IsAddInFuncs;
        }
        public SupBookRecord(String url, String[] sheetNames)
        {
            field_1_number_of_sheets = (short)sheetNames.Length;
            field_2_encoded_url = url;
            field_3_sheet_names = sheetNames;
            _isAddInFunctions = false;
        }

        /**
         * Constructs a Extern Sheet record and Sets its fields appropriately.
         *
         * @param id     id must be 0x16 or an exception will be throw upon validation
         * @param size  the size of the data area of the record
         * @param data  data of the record (should not contain sid/len)
         */
        public SupBookRecord(RecordInputStream in1)
        {
            int recLen = in1.Remaining;

            field_1_number_of_sheets = in1.ReadShort();

            if (recLen > SMALL_RECORD_SIZE)
            {
                // 5.38.1 External References
                _isAddInFunctions = false;

                field_2_encoded_url = in1.ReadString();
                String[] sheetNames = new String[field_1_number_of_sheets];
                for (int i = 0; i < sheetNames.Length; i++)
                {
                    sheetNames[i] = in1.ReadString();
                }
                field_3_sheet_names = sheetNames;
                return;
            }
            // else not 'External References'
            field_2_encoded_url = null;
            field_3_sheet_names = null;

            short nextShort = in1.ReadShort();
            if (nextShort == TAG_INTERNAL_REFERENCES)
            {
                // 5.38.2 'Internal References'
                _isAddInFunctions = false;
            }
            else if (nextShort == TAG_ADD_IN_FUNCTIONS)
            {
                // 5.38.3 'Add-In Functions'
                _isAddInFunctions = true;
                if (field_1_number_of_sheets != 1)
                {
                    throw new Exception("Expected 0x0001 for number of sheets field in 'Add-In Functions' but got ("
                         + field_1_number_of_sheets + ")");
                }
            }
            else
            {
                throw new Exception("invalid EXTERNALBOOK code ("
                         + StringUtil.ToHexString(nextShort) + ")");
            }
        }

        public bool IsExternalReferences
        {
            get { return field_3_sheet_names != null; }
        }
        public bool IsInternalReferences
        {
            get
            {
                return field_3_sheet_names == null && !_isAddInFunctions;
            }
        }
        public bool IsAddInFunctions
        {
            get
            {
                return field_3_sheet_names == null && _isAddInFunctions;
            }
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(GetType().Name).Append(" [SUPBOOK ");

            if (IsExternalReferences)
            {
                sb.Append("External References");
                sb.Append(" nSheets=").Append(field_1_number_of_sheets);
                sb.Append(" url=").Append(field_2_encoded_url);
            }
            else if (_isAddInFunctions)
            {
                sb.Append("Add-In Functions");
            }
            else
            {
                sb.Append("Internal References ");
                sb.Append(" nSheets= ").Append(field_1_number_of_sheets);
            }
            return sb.ToString();
        }
        protected override int DataSize
        {
            get
            {
                if (!IsExternalReferences)
                {
                    return SMALL_RECORD_SIZE;
                }
                int sum = 2; // u16 number of sheets

                sum += StringUtil.GetEncodedSize(field_2_encoded_url);

                for (int i = 0; i < field_3_sheet_names.Length; i++)
                {
                    sum += StringUtil.GetEncodedSize(field_3_sheet_names[i]);
                }
                return sum;
            }
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_number_of_sheets);

            if (IsExternalReferences)
            {
                StringUtil.WriteUnicodeString(out1, field_2_encoded_url);

                for (int i = 0; i < field_3_sheet_names.Length; i++)
                {
                    StringUtil.WriteUnicodeString(out1, field_3_sheet_names[i]);
                }
            }
            else
            {
                int field2val = _isAddInFunctions ? TAG_ADD_IN_FUNCTIONS : TAG_INTERNAL_REFERENCES;

                out1.WriteShort(field2val);
            }
        }

        public short NumberOfSheets
        {
            get { return field_1_number_of_sheets; }
            set { field_1_number_of_sheets = value; }
        }

        public override short Sid
        {
            get { return sid; }
        }
        public String URL
        {
            get
            {
                String encodedUrl = field_2_encoded_url;
                switch ((int)encodedUrl[0])
                {
                    case 0: // Reference to an empty workbook name
                        return encodedUrl.Substring(1); // will this just be empty string?
                    case 1: // encoded file name
                        return DecodeFileName(encodedUrl);
                    case 2: // Self-referential external reference
                        return encodedUrl.Substring(1);

                }
                return encodedUrl;
            }
        }
        private static String DecodeFileName(String encodedUrl)
        {
            return encodedUrl.Substring(1);
            // TODO the following special characters may appear in the rest of the string, and need to get interpreted
            /* see "MICROSOFT OFFICE EXCEL 97-2007  BINARY FILE FORMAT SPECIFICATION"
            chVolume  1 
            chSameVolume  2 
            chDownDir  3
            chUpDir  4 
            chLongVolume  5
            chStartupDir  6
            chAltStartupDir 7
            chLibDir  8
            */
        }
        public String[] SheetNames
        {
            get
            {
                return (String[])field_3_sheet_names.Clone();
            }
        }
    }
}