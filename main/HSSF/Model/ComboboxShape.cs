namespace NPOI.HSSF.Model
{
    using System;
    using NPOI.DDF;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Record;
    using NPOI.SS.UserModel;

    [Obsolete]
    public class ComboboxShape:AbstractShape
    {
        private EscherContainerRecord spContainer;
        private ObjRecord objRecord;

        /**
         * Creates the low evel records for a combobox.
         *
         * @param hssfShape The highlevel shape.
         * @param shapeId   The shape id to use for this shape.
         */
        public ComboboxShape(HSSFSimpleShape hssfShape, int shapeId)
        {
            spContainer = CreateSpContainer(hssfShape, shapeId);
            objRecord = CreateObjRecord(hssfShape, shapeId);
        }

        /**
         * Creates the low level OBJ record for this shape.
         */
        private ObjRecord CreateObjRecord(HSSFSimpleShape shape, int shapeId)
        {
            ObjRecord obj = new ObjRecord();
            CommonObjectDataSubRecord c = new CommonObjectDataSubRecord();
            c.ObjectType = CommonObjectType.ComboBox;
            c.ObjectId = shapeId;
            c.IsLocked = true;
            c.IsPrintable = false;
            c.IsAutoFill = true;
            c.IsAutoline = false;

            FtCblsSubRecord f = new FtCblsSubRecord();

            LbsDataSubRecord l = LbsDataSubRecord.CreateAutoFilterInstance();

            EndSubRecord e = new EndSubRecord();

            obj.AddSubRecord(c);
            obj.AddSubRecord(f);
            obj.AddSubRecord(l);
            obj.AddSubRecord(e);

            return obj;
        }

        /**
         * Generates the escher shape records for this shape.
         */
        private EscherContainerRecord CreateSpContainer(HSSFSimpleShape shape, int shapeId)
        {
            EscherContainerRecord spContainer = new EscherContainerRecord();
            EscherSpRecord sp = new EscherSpRecord();
            EscherOptRecord opt = new EscherOptRecord();
            EscherClientDataRecord clientData = new EscherClientDataRecord();

            spContainer.RecordId=(EscherContainerRecord.SP_CONTAINER);
            spContainer.Options=((short)0x000F);
            sp.RecordId=(EscherSpRecord.RECORD_ID);
            sp.Options=((short)((EscherAggregate.ST_HOSTCONTROL << 4) | 0x2));

            sp.ShapeId=(shapeId);
            sp.Flags=(EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE);
            opt.RecordId=(EscherOptRecord.RECORD_ID);
            opt.AddEscherProperty(new EscherBoolProperty(EscherProperties.PROTECTION__LOCKAGAINSTGROUPING, 17039620));
            opt.AddEscherProperty(new EscherBoolProperty(EscherProperties.TEXT__SIZE_TEXT_TO_FIT_SHAPE, 0x00080008));
            opt.AddEscherProperty(new EscherBoolProperty(EscherProperties.LINESTYLE__NOLINEDRAWDASH, 0x00080000));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.GROUPSHAPE__PRINT, 0x00020000));

            HSSFClientAnchor userAnchor = (HSSFClientAnchor)shape.Anchor;
            userAnchor.AnchorType = (AnchorType)1;
            EscherRecord anchor = CreateAnchor(userAnchor);
            clientData.RecordId=(EscherClientDataRecord.RECORD_ID);
            clientData.Options=((short)0x0000);

            spContainer.AddChildRecord(sp);
            spContainer.AddChildRecord(opt);
            spContainer.AddChildRecord(anchor);
            spContainer.AddChildRecord(clientData);

            return spContainer;
        }

        public override EscherContainerRecord SpContainer
        {
            get
            {
                return spContainer;
            }
        }

        public override ObjRecord ObjRecord
        {
            get
            {
                return objRecord;
            }
        }
    }
}
