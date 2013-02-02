/* ====================================================================
   Copyright 2002-2004   Apache Software Foundation

   Licensed Under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


namespace NPOI.HSSF.UserModel
{
    using System;
    using System.IO;
    using System.Collections;

    using NPOI.HSSF.Record;
    using NPOI.Util;
    using NPOI.POIFS.FileSystem;
    using NPOI.DDF;


    /**
     * Represents binary object (i.e. OLE) data stored in the file.  Eg. A GIF, JPEG etc...
     *
     * @author Daniel Noll
     */
    public class HSSFObjectData : HSSFPicture
    {

        /**
         * Reference to the filesystem root, required for retrieving the object data.
         */
        private DirectoryEntry _root;


        public HSSFObjectData(EscherContainerRecord spContainer, ObjRecord objRecord, DirectoryEntry _root)
            : base(spContainer, objRecord)
        {
            this._root = _root;
        }
        /**
         * Returns the OLE2 Class Name of the object
         */
        public String OLE2ClassName
        {
            get
            {
                return FindObjectRecord().OLEClassName;
            }
        }

        /**
         * Gets the object data. Only call for ones that have
         *  data though. See {@link #hasDirectoryEntry()}
         *
         * @return the object data as an OLE2 directory.
         * @ if there was an error Reading the data.
         */
        public DirectoryEntry GetDirectory()
        {
            EmbeddedObjectRefSubRecord subRecord = FindObjectRecord();

            int? streamId = ((EmbeddedObjectRefSubRecord)subRecord).StreamId;
            String streamName = "MBD" + HexDump.ToHex((int)streamId);

            Entry entry = _root.GetEntry(streamName);
            if (entry is DirectoryEntry)
            {
                return (DirectoryEntry)entry;
            }
            else
            {
                throw new IOException("Stream " + streamName + " was not an OLE2 directory");
            }
        }

        /**
         * Returns the data portion, for an ObjectData
         *  that doesn't have an associated POIFS Directory
         *  Entry
         */
        public byte[] GetObjectData()
        {
            return FindObjectRecord().ObjectData;
        }

        /**
         * Does this ObjectData have an associated POIFS 
         *  Directory Entry?
         * (Not all do, those that don't have a data portion)
         */
        public bool HasDirectoryEntry()
        {
            EmbeddedObjectRefSubRecord subRecord = FindObjectRecord();

            // 'stream id' field tells you
            int? streamId = subRecord.StreamId;
            return streamId != null && streamId != 0;
        }

        /**
         * Finds the EmbeddedObjectRefSubRecord, or throws an 
         *  Exception if there wasn't one
         */
        public EmbeddedObjectRefSubRecord FindObjectRecord()
        {
            IEnumerator subRecordIter = GetObjRecord().SubRecords.GetEnumerator();

            while (subRecordIter.MoveNext())
            {
                Object subRecord = subRecordIter.Current;
                if (subRecord is EmbeddedObjectRefSubRecord)
                {
                    return (EmbeddedObjectRefSubRecord)subRecord;
                }
            }

            throw new InvalidOperationException("Object data does not contain a reference to an embedded object OLE2 directory");
        }


        protected override EscherContainerRecord CreateSpContainer()
        {
            throw new InvalidOperationException("HSSFObjectData cannot be created from scratch");
        }


        protected override ObjRecord CreateObjRecord()
        {
            throw new InvalidOperationException("HSSFObjectData cannot be created from scratch");
        }


        internal override void AfterRemove(HSSFPatriarch patriarch)
        {
            throw new InvalidOperationException("HSSFObjectData cannot be created from scratch");
        }


        internal override void AfterInsert(HSSFPatriarch patriarch)
        {
            EscherAggregate agg = patriarch.GetBoundAggregate();
            agg.AssociateShapeToObjRecord(GetEscherContainer().GetChildById(EscherClientDataRecord.RECORD_ID), GetObjRecord());
            EscherBSERecord bse =
                    ((HSSFWorkbook)patriarch.Sheet.Workbook).Workbook.GetBSERecord(PictureIndex);
            bse.Ref = (bse.Ref + 1);
        }

        internal override HSSFShape CloneShape()
        {
            EscherContainerRecord spContainer = new EscherContainerRecord();
            byte[] inSp = GetEscherContainer().Serialize();
            spContainer.FillFields(inSp, 0, new DefaultEscherRecordFactory());
            ObjRecord obj = (ObjRecord)GetObjRecord().CloneViaReserialise();
            return new HSSFObjectData(spContainer, obj, _root);
        }
    }
}