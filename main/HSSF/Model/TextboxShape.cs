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
    using NPOI.DDF;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Record;

    /// <summary>
    /// Represents an textbox shape and Converts between the highlevel records
    /// and lowlevel records for an oval.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    [Obsolete]
    public class TextboxShape : AbstractShape
    {
        private EscherContainerRecord spContainer;
        private TextObjectRecord textObjectRecord;
        private ObjRecord objRecord;
        private EscherTextboxRecord escherTextbox;

        /// <summary>
        /// Creates the low evel records for a textbox.
        /// </summary>
        /// <param name="hssfShape">The highlevel shape.</param>
        /// <param name="shapeId">The shape id to use for this shape.</param>
        public TextboxShape(HSSFTextbox hssfShape, int shapeId)
        {
            spContainer = CreateSpContainer(hssfShape, shapeId);
            objRecord = CreateObjRecord(hssfShape, shapeId);
            textObjectRecord = CreateTextObjectRecord(hssfShape, shapeId);
        }

        /// <summary>
        /// Creates the lowerlevel OBJ records for this shape.
        /// </summary>
        /// <param name="hssfShape">The HSSF shape.</param>
        /// <param name="shapeId">The shape id.</param>
        /// <returns></returns>
        private ObjRecord CreateObjRecord(HSSFTextbox hssfShape, int shapeId)
        {
            HSSFShape shape = hssfShape;

            ObjRecord obj = new ObjRecord();
            CommonObjectDataSubRecord c = new CommonObjectDataSubRecord();
            c.ObjectType = (CommonObjectType)((HSSFSimpleShape)shape).ShapeType;
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
        /// Creates the lowerlevel escher records for this shape.
        /// </summary>
        /// <param name="hssfShape">The HSSF shape.</param>
        /// <param name="shapeId">The shape id.</param>
        /// <returns></returns>
        private EscherContainerRecord CreateSpContainer(HSSFTextbox hssfShape, int shapeId)
        {
            HSSFTextbox shape = hssfShape;

            EscherContainerRecord spContainer = new EscherContainerRecord();
            EscherSpRecord sp = new EscherSpRecord();
            EscherOptRecord opt = new EscherOptRecord();
            EscherRecord anchor = new EscherClientAnchorRecord();
            EscherClientDataRecord clientData = new EscherClientDataRecord();
            escherTextbox = new EscherTextboxRecord();

            spContainer.RecordId=EscherContainerRecord.SP_CONTAINER;
            spContainer.Options=(short)0x000F;
            sp.RecordId=EscherSpRecord.RECORD_ID;
            sp.Options=(short)((EscherAggregate.ST_TEXTBOX << 4) | 0x2);

            sp.ShapeId=shapeId;
            sp.Flags=EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE;
            opt.RecordId=EscherOptRecord.RECORD_ID;
            //        opt.AddEscherProperty( new EscherBoolProperty( EscherProperties.PROTECTION__LOCKAGAINSTGROUPING, 262144 ) );
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TEXT__TEXTID, 0));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TEXT__TEXTLEFT, shape.MarginLeft));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TEXT__TEXTRIGHT, shape.MarginRight));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TEXT__TEXTBOTTOM, shape.MarginBottom));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TEXT__TEXTTOP, shape.MarginTop));

            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TEXT__WRAPTEXT, 0));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TEXT__ANCHORTEXT, 0));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.GROUPSHAPE__PRINT, 0x00080000));

            AddStandardOptions(shape, opt);
            HSSFAnchor userAnchor = shape.Anchor;
            //        if (userAnchor.IsHorizontallyFlipped())
            //            sp.Flags(sp.Flags | EscherSpRecord.FLAG_FLIPHORIZ);
            //        if (userAnchor.IsVerticallyFlipped())
            //            sp.Flags(sp.Flags | EscherSpRecord.FLAG_FLIPVERT);
            anchor = CreateAnchor(userAnchor);
            clientData.RecordId=EscherClientDataRecord.RECORD_ID;
            clientData.Options=(short)0x0000;
            escherTextbox.RecordId=EscherTextboxRecord.RECORD_ID;
            escherTextbox.Options=(short)0x0000;

            spContainer.AddChildRecord(sp);
            spContainer.AddChildRecord(opt);
            spContainer.AddChildRecord(anchor);
            spContainer.AddChildRecord(clientData);
            spContainer.AddChildRecord(escherTextbox);

            return spContainer;
        }

        /// <summary>
        /// Textboxes also have an extra TXO record associated with them that most
        /// other shapes dont have.
        /// </summary>
        /// <param name="hssfShape">The HSSF shape.</param>
        /// <param name="shapeId">The shape id.</param>
        /// <returns></returns>
        private TextObjectRecord CreateTextObjectRecord(HSSFTextbox hssfShape, int shapeId)
        {
            HSSFTextbox shape = hssfShape;

            TextObjectRecord obj = new TextObjectRecord();
            obj.HorizontalTextAlignment=hssfShape.HorizontalAlignment;
            obj.VerticalTextAlignment=hssfShape.VerticalAlignment;
            obj.IsTextLocked=true;
            obj.TextOrientation=TextOrientation.None;
            int frLength = (shape.String.NumFormattingRuns + 1) * 8;
            obj.Str=shape.String;

            return obj;
        }
        /// <summary>
        /// The shape container and it's children that can represent this
        /// shape.
        /// </summary>
        /// <value></value>
        public override EscherContainerRecord SpContainer
        {
            get { return spContainer; }
        }
        /// <summary>
        /// The object record that is associated with this shape.
        /// </summary>
        /// <value></value>
        public override ObjRecord ObjRecord
        {
            get { return objRecord; }
        }
        /// <summary>
        /// The TextObject record that is associated with this shape.
        /// </summary>
        /// <value></value>
        public TextObjectRecord TextObjectRecord
        {
            get{return textObjectRecord;}
        }

        /// <summary>
        /// Gets the EscherTextbox record.
        /// </summary>
        /// <value>The EscherTextbox record.</value>
        public EscherRecord EscherTextbox
        {
            get{return escherTextbox;}
        }
    }
}
