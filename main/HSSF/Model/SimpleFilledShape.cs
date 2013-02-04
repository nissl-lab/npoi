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
    using NPOI.DDF;
    using NPOI.HSSF.UserModel;

    [Obsolete]
    public class SimpleFilledShape: AbstractShape
    {
        private EscherContainerRecord spContainer;
        private ObjRecord objRecord;

        /// <summary>
        /// Creates the low evel records for an oval.
        /// </summary>
        /// <param name="hssfShape">The highlevel shape.</param>
        /// <param name="shapeId">The shape id to use for this shape.</param>
        public SimpleFilledShape(HSSFSimpleShape hssfShape, int shapeId)
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
        private EscherContainerRecord CreateSpContainer(HSSFSimpleShape hssfShape, int shapeId)
        {
            HSSFShape shape = hssfShape;

            EscherContainerRecord spContainer = new EscherContainerRecord();
            EscherSpRecord sp = new EscherSpRecord();
            EscherOptRecord opt = new EscherOptRecord();
            EscherClientDataRecord clientData = new EscherClientDataRecord();

            spContainer.RecordId=EscherContainerRecord.SP_CONTAINER;
            spContainer.Options=(short)0x000F;
            sp.RecordId=EscherSpRecord.RECORD_ID;
            short shapeType = objTypeToShapeType(hssfShape.ShapeType);
            sp.Options=(short)((shapeType << 4) | 0x2);
            sp.ShapeId=shapeId;
            sp.Flags=EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE;
            opt.RecordId=EscherOptRecord.RECORD_ID;
            AddStandardOptions(shape, opt);
            EscherRecord anchor = CreateAnchor(shape.Anchor);
            clientData.RecordId=EscherClientDataRecord.RECORD_ID;
            clientData.Options=(short)0x0000;

            spContainer.AddChildRecord(sp);
            spContainer.AddChildRecord(opt);
            spContainer.AddChildRecord(anchor);
            spContainer.AddChildRecord(clientData);

            return spContainer;
        }

        private short objTypeToShapeType(int objType)
        {
            short shapeType;
            if (objType == HSSFSimpleShape.OBJECT_TYPE_OVAL)
                shapeType = EscherAggregate.ST_ELLIPSE;
            else if (objType == HSSFSimpleShape.OBJECT_TYPE_RECTANGLE)
                shapeType = EscherAggregate.ST_RECTANGLE;
            else
                throw new ArgumentException("Unable to handle an object of this type");
            return shapeType;
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
            c.ObjectType = (CommonObjectType)((HSSFSimpleShape)shape).ShapeType;
            c.ObjectId = shapeId;
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
            get
            {
                return spContainer;
            }
        }
        /// <summary>
        /// The object record that is associated with this shape.
        /// </summary>
        /// <value></value>
        public override ObjRecord ObjRecord
        {
            get
            {
                return objRecord;
            }
        }

    }
}