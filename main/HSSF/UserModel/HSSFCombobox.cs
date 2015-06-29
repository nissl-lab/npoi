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
using NPOI.DDF;
using NPOI.HSSF.Record;
using NPOI.SS.UserModel;

namespace NPOI.HSSF.UserModel
{
    public class HSSFCombobox : HSSFSimpleShape
    {
        public HSSFCombobox(EscherContainerRecord spContainer, ObjRecord objRecord)
            : base(spContainer, objRecord)
        {

        }

        public HSSFCombobox(HSSFShape parent, HSSFAnchor anchor)
            : base(parent, anchor)
        {

            base.ShapeType = (OBJECT_TYPE_COMBO_BOX);
            CommonObjectDataSubRecord cod = (CommonObjectDataSubRecord)GetObjRecord().SubRecords[0];
            cod.ObjectType = CommonObjectType.ComboBox;
        }

        protected override TextObjectRecord CreateTextObjRecord()
        {
            return null;
        }

        protected override EscherContainerRecord CreateSpContainer()
        {
            EscherContainerRecord spContainer = new EscherContainerRecord();
            EscherSpRecord sp = new EscherSpRecord();
            EscherOptRecord opt = new EscherOptRecord();
            EscherClientDataRecord clientData = new EscherClientDataRecord();

            spContainer.RecordId = (EscherContainerRecord.SP_CONTAINER);
            spContainer.Options = ((short)0x000F);
            sp.RecordId = (EscherSpRecord.RECORD_ID);
            sp.Options = ((short)((EscherAggregate.ST_HOSTCONTROL << 4) | 0x2));

            sp.Flags = (EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE);
            opt.RecordId = (EscherOptRecord.RECORD_ID);
            opt.AddEscherProperty(new EscherBoolProperty(EscherProperties.PROTECTION__LOCKAGAINSTGROUPING, 17039620));
            opt.AddEscherProperty(new EscherBoolProperty(EscherProperties.TEXT__SIZE_TEXT_TO_FIT_SHAPE, 0x00080008));
            opt.AddEscherProperty(new EscherBoolProperty(EscherProperties.LINESTYLE__NOLINEDRAWDASH, 0x00080000));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.GROUPSHAPE__PRINT, 0x00020000));

            HSSFClientAnchor userAnchor = (HSSFClientAnchor)Anchor;
            userAnchor.AnchorType = (AnchorType)(1);
            EscherRecord anchor = userAnchor.GetEscherAnchor();
            clientData.RecordId = (EscherClientDataRecord.RECORD_ID);
            clientData.Options = ((short)0x0000);

            spContainer.AddChildRecord(sp);
            spContainer.AddChildRecord(opt);
            spContainer.AddChildRecord(anchor);
            spContainer.AddChildRecord(clientData);

            return spContainer;
        }

        protected override ObjRecord CreateObjRecord()
        {
            ObjRecord obj = new ObjRecord();
            CommonObjectDataSubRecord c = new CommonObjectDataSubRecord();
            c.ObjectType = CommonObjectType.ComboBox;
            c.IsLocked = (true);
            c.IsPrintable = (false);
            c.IsAutoFill = (true);
            c.IsAutoline = (false);
            FtCblsSubRecord f = new FtCblsSubRecord();
            LbsDataSubRecord l = LbsDataSubRecord.CreateAutoFilterInstance();
            EndSubRecord e = new EndSubRecord();
            obj.AddSubRecord(c);
            obj.AddSubRecord(f);
            obj.AddSubRecord(l);
            obj.AddSubRecord(e);
            return obj;
        }

        public override int ShapeType
        {
            get { return base.ShapeType; }
            set { throw new InvalidOperationException("Shape type can not be changed in " + this.GetType().Name); }
        }
    }
}
