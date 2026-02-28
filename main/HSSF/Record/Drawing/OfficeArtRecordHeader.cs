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
using System;
using System.Collections.Generic;
using System.Text;
using NPOI.Util;

namespace NPOI.HSSF.Record.Drawing
{
    /// <summary>
    /// specifies the common record header for all the OfficeArt records.
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    public class OfficeArtRecordHeader
    {
        private short field_1_recVer_Instance;
        private ushort field_2_recType;
        private int field_3_recLen;

        private BitField recVer = BitFieldFactory.GetInstance(0xF);
        private BitField recInstance = BitFieldFactory.GetInstance(0xFFF0);

        public OfficeArtRecordHeader()
        {

        }
        public OfficeArtRecordHeader(RecordInputStream ris)
        {
            field_1_recVer_Instance = ris.ReadShort();
            field_2_recType = (ushort)ris.ReadUShort();
            field_3_recLen = ris.ReadInt();
        }
        public int DataSize
        {
            get { return 8; }
        }

        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_recVer_Instance);
            out1.WriteShort(field_2_recType);
            out1.WriteInt(field_3_recLen);
        }
        /// <summary>
        /// specifies the version if the record is an atom. If the record is a container, this field MUST contain 0xF.
        /// </summary>
        public short Ver
        {
            get { return recVer.GetShortValue(field_1_recVer_Instance); }
            set { field_1_recVer_Instance = recVer.SetShortValue(field_1_recVer_Instance, value); }
        }
        /// <summary>
        /// An unsigned integer that differentiates an atom from the other atoms that are contained in the record.
        /// </summary>
        public short Instance
        {
            get { return recInstance.GetShortValue(field_1_recVer_Instance); }
            set { field_1_recVer_Instance = recInstance.SetShortValue(field_1_recVer_Instance, value); }
        }
        /// <summary>
        /// specifies the type of the record. This value MUST be from 0xF000 through 0xFFFF, inclusive.
        /// </summary>
        public ushort Type
        {
            get { return field_2_recType; }
            set { field_2_recType = value; }
        }
        /// <summary>
        /// that specifies the length, in bytes, of the record. 
        /// If the record is an atom, this value specifies the length of the atom, excluding the header. 
        /// If the record is a container, this value specifies the sum of the lengths of the atoms that 
        /// the record contains, plus the length of the record header for each atom.
        /// </summary>
        public int Len
        {
            get { return field_3_recLen; }
            set { field_3_recLen = value; }
        }
    }
}

