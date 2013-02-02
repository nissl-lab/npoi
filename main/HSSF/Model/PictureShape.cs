
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
    using NPOI.HSSF.Record;
    using NPOI.DDF;
    using NPOI.HSSF.UserModel;


    /// <summary>
    /// Represents a picture shape and Creates all specific low level records.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class PictureShape: AbstractShape
    {
        private EscherContainerRecord spContainer;
        private ObjRecord objRecord;

        /// <summary>
        /// Creates the line shape from the highlevel user shape.  All low level
        /// records are Created at this point.
        /// </summary>
        /// <param name="hssfShape">The user model shape.</param>
        /// <param name="shapeId">The identifier to use for this shape.</param>
        public PictureShape(HSSFSimpleShape hssfShape, int shapeId)
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
            HSSFPicture shape = (HSSFPicture)hssfShape;

            EscherContainerRecord spContainer = new EscherContainerRecord();
            EscherSpRecord sp = new EscherSpRecord();
            EscherOptRecord opt = new EscherOptRecord();
            EscherRecord anchor;
            EscherClientDataRecord clientData = new EscherClientDataRecord();

            spContainer.RecordId=EscherContainerRecord.SP_CONTAINER;
            spContainer.Options=(short)0x000F;
            sp.RecordId=EscherSpRecord.RECORD_ID;
            sp.Options=(short)((EscherAggregate.ST_PICTUREFRAME << 4) | 0x2);

            sp.ShapeId=shapeId;
            sp.Flags=EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE;
            opt.RecordId=EscherOptRecord.RECORD_ID;
            //        opt.AddEscherProperty( new EscherBoolProperty( EscherProperties.PROTECTION__LOCKAGAINSTGROUPING, 0x00800080 ) ;
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.BLIP__BLIPTODISPLAY, false, true, shape.PictureIndex));
            //        opt.AddEscherProperty( new EscherComplexProperty( EscherProperties.BLIP__BLIPFILENAME, true, new byte[] { (byte)0x74, (byte)0x00, (byte)0x65, (byte)0x00, (byte)0x73, (byte)0x00, (byte)0x74, (byte)0x00, (byte)0x00, (byte)0x00 } ) ;
            //        opt.AddEscherProperty( new EscherSimpleProperty( EscherProperties.Fill__FillTYPE, 0x00000003 ) ;
            AddStandardOptions(shape, opt);
            HSSFAnchor userAnchor = shape.Anchor;
            if (userAnchor.IsHorizontallyFlipped)
                sp.Flags=sp.Flags | EscherSpRecord.FLAG_FLIPHORIZ;
            if (userAnchor.IsVerticallyFlipped)
                sp.Flags=sp.Flags | EscherSpRecord.FLAG_FLIPVERT;
            anchor = CreateAnchor(userAnchor);
            clientData.RecordId=EscherClientDataRecord.RECORD_ID;
            clientData.Options=(short)0x0000;

            spContainer.AddChildRecord(sp);
            spContainer.AddChildRecord(opt);
            spContainer.AddChildRecord(anchor);
            spContainer.AddChildRecord(clientData);

            return spContainer;
        }

        /// <summary>
        /// Creates the low level OBJ record for this shape.
        /// </summary>
        /// <param name="hssfShape">The HSSFShape.</param>
        /// <param name="shapeId">The shape id.</param>
        /// <returns></returns>
        private ObjRecord CreateObjRecord(HSSFShape hssfShape, int shapeId)
        {
            HSSFShape shape = hssfShape;

            ObjRecord obj = new ObjRecord();
            CommonObjectDataSubRecord c = new CommonObjectDataSubRecord();
            c.ObjectType=(CommonObjectType)((HSSFSimpleShape)shape).ShapeType;
            //        c.ObjectId((short) ( 1 ));
            c.ObjectId = GetCmoObjectId(shapeId);
            c.IsLocked=true;
            c.IsPrintable=true;
            c.IsAutoFill=true;
            c.IsAutoline=true;
            //        c.Reserved2( 0x012C0A84 );
            c.Reserved2=(0x0);
            //        UnknownRecord sub1 = new UnknownRecord( (short)0x7, (short)0x2, new byte[] { 0x09, 0x00 } );
            //        UnknownRecord sub2 = new UnknownRecord( (short)0x8, (short)0x2, new byte[] { 0x01, 0x00 } );
            EndSubRecord e = new EndSubRecord();

            obj.AddSubRecord(c);
            //        obj.AddSubRecord( sub1 );
            //        obj.AddSubRecord( sub2 );
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