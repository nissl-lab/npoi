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

using System.Collections.Generic;
using NPOI.DDF;
using NPOI.HSSF.Record;
using NPOI.POIFS.FileSystem;

namespace NPOI.HSSF.UserModel
{
    /**
     * @author Evgeniy Berlog
     * date: 05.06.12
     */
    public class HSSFShapeFactory
    {
        private const short OBJECT_TYPE_LINE = 1;
        private const short OBJECT_TYPE_RECTANGLE = 2;
        private const short OBJECT_TYPE_OVAL = 3;
        private const short OBJECT_TYPE_ARC = 4;
        private const short OBJECT_TYPE_PICTURE = 8;

        /**
         * build shape tree from escher container
         * @param container root escher container from which escher records must be taken
         * @param agg - EscherAggregate
         * @param out - shape container to which shapes must be added
         * @param root - node to create HSSFObjectData shapes
         */
        public static void CreateShapeTree(EscherContainerRecord container, EscherAggregate agg,
            HSSFShapeContainer out1, DirectoryNode root)
        {
            if (container.RecordId == EscherContainerRecord.SPGR_CONTAINER)
            {
                ObjRecord obj = null;
                EscherClientDataRecord clientData = (EscherClientDataRecord)((EscherContainerRecord)container.GetChild(0)).GetChildById(EscherClientDataRecord.RECORD_ID);
                if (null != clientData)
                {
                    obj = (ObjRecord)agg.GetShapeToObjMapping()[clientData];
                }
                HSSFShapeGroup group = new HSSFShapeGroup(container, obj);
                IList<EscherContainerRecord> children = container.ChildContainers;
                // skip the first child record, it is group descriptor
                for (int i = 0; i < children.Count; i++)
                {
                    EscherContainerRecord spContainer = children[(i)];
                    if (i != 0)
                    {
                        CreateShapeTree(spContainer, agg, group, root);
                    }
                }
                out1.AddShape(group);
            }
            else if (container.RecordId == EscherContainerRecord.SP_CONTAINER)
            {
                Dictionary<EscherRecord, Record.Record> shapeToObj = agg.GetShapeToObjMapping();
                ObjRecord objRecord = null;
                TextObjectRecord txtRecord = null;

                foreach (EscherRecord record in container.ChildRecords)
                {
                    switch (record.RecordId)
                    {
                        case EscherClientDataRecord.RECORD_ID:
                            objRecord = (ObjRecord)shapeToObj[(record)];
                            break;
                        case EscherTextboxRecord.RECORD_ID:
                            txtRecord = (TextObjectRecord)shapeToObj[(record)];
                            break;
                    }
                }
                if (IsEmbeddedObject(objRecord))
                {
                    HSSFObjectData objectData = new HSSFObjectData(container, objRecord, root);
                    out1.AddShape(objectData);
                    return;
                }
                CommonObjectDataSubRecord cmo = (CommonObjectDataSubRecord)objRecord.SubRecords[0];
                HSSFShape shape;
                switch (cmo.ObjectType)
                {
                    case CommonObjectType.Picture:
                        shape = new HSSFPicture(container, objRecord);
                        break;
                    case CommonObjectType.Rectangle:
                        shape = new HSSFSimpleShape(container, objRecord, txtRecord);
                        break;
                    case CommonObjectType.Line:
                        shape = new HSSFSimpleShape(container, objRecord);
                        break;
                    case CommonObjectType.ComboBox:
                        shape = new HSSFCombobox(container, objRecord);
                        break;
                    case CommonObjectType.MicrosoftOfficeDrawing:
                        EscherOptRecord optRecord = (EscherOptRecord)container.GetChildById(EscherOptRecord.RECORD_ID);
                        EscherProperty property = optRecord.Lookup(EscherProperties.GEOMETRY__VERTICES);
                        if (null != property)
                        {
                            shape = new HSSFPolygon(container, objRecord, txtRecord);
                        }
                        else
                        {
                            shape = new HSSFSimpleShape(container, objRecord, txtRecord);
                        }
                        break;
                    case CommonObjectType.Text:
                        shape = new HSSFTextbox(container, objRecord, txtRecord);
                        break;
                    case CommonObjectType.Comment:
                        shape = new HSSFComment(container, objRecord, txtRecord, agg.GetNoteRecordByObj(objRecord));
                        break;
                    default:
                        shape = new HSSFSimpleShape(container, objRecord, txtRecord);
                        break;
                }
                out1.AddShape(shape);
            }
        }

        private static bool IsEmbeddedObject(ObjRecord obj)
        {
            //Iterator<SubRecord> subRecordIter = obj.SubRecords;
            foreach (SubRecord sub in obj.SubRecords)
            {
                //SubRecord sub = subRecordIter.next();
                if (sub is EmbeddedObjectRefSubRecord)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
