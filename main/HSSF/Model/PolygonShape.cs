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


namespace NPOI.HSSF.Model
{
    using System;

    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using NPOI.DDF;
    using NPOI.Util;

    [Obsolete]
    public class PolygonShape: AbstractShape
    {
        public const short OBJECT_TYPE_MICROSOFT_OFFICE_DRAWING = 30;
        private EscherContainerRecord spContainer;
        private ObjRecord objRecord;

        /// <summary>
        /// Creates the low evel records for an polygon.
        /// </summary>
        /// <param name="hssfShape">The highlevel shape.</param>
        /// <param name="shapeId">The shape id to use for this shape.</param>
        public PolygonShape(HSSFPolygon hssfShape, int shapeId)
        {
            spContainer = CreateSpContainer(hssfShape, shapeId);
            objRecord = CreateObjRecord(hssfShape, shapeId);
        }

        /// <summary>
        /// Creates the lowerlevel escher records for this shape.
        /// </summary>
        /// <param name="hssfShape">The HSSF shape.</param>
        /// <param name="shapeId">The shape id.</param>
        /// <returns></returns>
        private EscherContainerRecord CreateSpContainer(HSSFPolygon hssfShape, int shapeId)
        {
            HSSFShape shape = hssfShape;

            EscherContainerRecord spContainer = new EscherContainerRecord();
            EscherSpRecord sp = new EscherSpRecord();
            EscherOptRecord opt = new EscherOptRecord();
            EscherClientDataRecord clientData = new EscherClientDataRecord();

            spContainer.RecordId=EscherContainerRecord.SP_CONTAINER;
            spContainer.Options=(short)0x000F;
            sp.RecordId=EscherSpRecord.RECORD_ID;
            sp.Options = (short)((EscherAggregate.ST_NOT_PRIMATIVE << 4) | 0x2);
            sp.ShapeId=shapeId;
            if (hssfShape.Parent == null)
                sp.Flags=EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE;
            else
                sp.Flags=EscherSpRecord.FLAG_CHILD | EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE;
            opt.RecordId=EscherOptRecord.RECORD_ID;
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TRANSFORM__ROTATION, false, false, 0));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.GEOMETRY__RIGHT, false, false, hssfShape.DrawAreaWidth));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.GEOMETRY__BOTTOM, false, false, hssfShape.DrawAreaHeight));
            opt.AddEscherProperty(new EscherShapePathProperty(EscherProperties.GEOMETRY__SHAPEPATH, EscherShapePathProperty.COMPLEX));
            EscherArrayProperty verticesProp = new EscherArrayProperty(EscherProperties.GEOMETRY__VERTICES, false, new byte[0]);
            verticesProp.NumberOfElementsInArray=(hssfShape.XPoints.Length + 1);
            verticesProp.NumberOfElementsInMemory=(hssfShape.XPoints.Length + 1);
            verticesProp.SizeOfElements=unchecked((short)0xFFF0);
            for (int i = 0; i < hssfShape.XPoints.Length; i++)
            {
                byte[] data = new byte[4];
                LittleEndian.PutShort(data, 0, (short)hssfShape.XPoints[i]);
                LittleEndian.PutShort(data, 2, (short)hssfShape.YPoints[i]);
                verticesProp.SetElement(i, data);
            }
            int point = hssfShape.XPoints.Length;
            byte[] data1 = new byte[4];
            LittleEndian.PutShort(data1, 0, (short)hssfShape.XPoints[0]);
            LittleEndian.PutShort(data1, 2, (short)hssfShape.YPoints[0]);
            verticesProp.SetElement(point, data1);
            opt.AddEscherProperty(verticesProp);
            EscherArrayProperty segmentsProp = new EscherArrayProperty(EscherProperties.GEOMETRY__SEGMENTINFO, false, null);
            segmentsProp.SizeOfElements=(0x0002);
            segmentsProp.NumberOfElementsInArray=(hssfShape.XPoints.Length * 2 + 4);
            segmentsProp.NumberOfElementsInMemory=(hssfShape.XPoints.Length * 2 + 4);
            segmentsProp.SetElement(0, new byte[] { (byte)0x00, (byte)0x40 });
            segmentsProp.SetElement(1, new byte[] { (byte)0x00, (byte)0xAC });
            for (int i = 0; i < hssfShape.XPoints.Length; i++)
            {
                segmentsProp.SetElement(2 + i * 2, new byte[] { (byte)0x01, (byte)0x00 });
                segmentsProp.SetElement(3 + i * 2, new byte[] { (byte)0x00, (byte)0xAC });
            }
            segmentsProp.SetElement(segmentsProp.NumberOfElementsInArray - 2, new byte[] { (byte)0x01, (byte)0x60 });
            segmentsProp.SetElement(segmentsProp.NumberOfElementsInArray - 1, new byte[] { (byte)0x00, (byte)0x80 });
            opt.AddEscherProperty(segmentsProp);
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.GEOMETRY__FILLOK, false, false, 0x00010001));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINESTARTARROWHEAD, false, false, 0x0));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINEENDARROWHEAD, false, false, 0x0));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINEENDCAPSTYLE, false, false, 0x0));

            AddStandardOptions(shape, opt);

            EscherRecord anchor = CreateAnchor(shape.Anchor);
            clientData.RecordId=(EscherClientDataRecord.RECORD_ID);
            clientData.Options=(short)0x0000;

            spContainer.AddChildRecord(sp);
            spContainer.AddChildRecord(opt);
            spContainer.AddChildRecord(anchor);
            spContainer.AddChildRecord(clientData);

            return spContainer;
        }

        /// <summary>
        /// Creates the lowerlevel OBJ records for this shape.
        /// </summary>
        /// <param name="hssfShape">The HSSF shape.</param>
        /// <param name="shapeId">The shape id.</param>
        /// <returns></returns>
        private ObjRecord CreateObjRecord(HSSFShape hssfShape, int shapeId)
        {
            HSSFShape shape = hssfShape;

            ObjRecord obj = new ObjRecord();
            CommonObjectDataSubRecord c = new CommonObjectDataSubRecord();
            c.ObjectType = (CommonObjectType)OBJECT_TYPE_MICROSOFT_OFFICE_DRAWING;
            c.ObjectId = GetCmoObjectId(shapeId);
            c.IsLocked = true;
            c.IsPrintable = true;
            c.IsAutoFill = true;
            c.IsAutoline = true;
            EndSubRecord e = new EndSubRecord();

            obj.AddSubRecord(c);
            obj.AddSubRecord(e);

            return obj;
        }
        /// <summary>
        /// The shape container and it's children that can represent this
        /// shape.
        /// </summary>
        /// <value></value>
        public override EscherContainerRecord SpContainer
        {
            get{return spContainer;}
        }
        /// <summary>
        /// The object record that is associated with this shape.
        /// </summary>
        /// <value></value>
        public override ObjRecord ObjRecord
        {
            get { return objRecord; }
        }

    }
}