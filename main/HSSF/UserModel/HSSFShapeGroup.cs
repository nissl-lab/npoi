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


namespace NPOI.HSSF.UserModel
{
    using NPOI.DDF;
    using NPOI.HSSF.Record;
    using System;
    using System.Collections;

    /// <summary>
    /// A shape Group may contain other shapes.  It was no actual form on the
    /// sheet.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class HSSFShapeGroup : HSSFShape, HSSFShapeContainer
    {
        IList shapes = new ArrayList();
        int x1 = 0;
        int y1 = 0;
        int x2 = 1023;
        int y2 = 255;
        private EscherSpgrRecord _spgrRecord;

        public HSSFShapeGroup(HSSFShape parent, HSSFAnchor anchor):base(parent, anchor)
        {
            
        }
        public HSSFShapeGroup(EscherContainerRecord spgrContainer, ObjRecord objRecord)
            : base(spgrContainer, objRecord)
        {
            // read internal and external coordinates from spgrContainer
            EscherContainerRecord spContainer = spgrContainer.ChildContainers[0];
            _spgrRecord = (EscherSpgrRecord)spContainer.GetChild(0);
            foreach (EscherRecord ch in spContainer.ChildRecords)
            {
                switch (ch.RecordId)
                {
                    case EscherSpgrRecord.RECORD_ID:
                        break;
                    case EscherClientAnchorRecord.RECORD_ID:
                        anchor = new HSSFClientAnchor((EscherClientAnchorRecord)ch);
                        break;
                    case EscherChildAnchorRecord.RECORD_ID:
                        anchor = new HSSFChildAnchor((EscherChildAnchorRecord)ch);
                        break;
                }
            }
        }
        /// <summary>
        /// Create another Group Under this Group.
        /// </summary>
        /// <param name="anchor">the position of the new Group.</param>
        /// <returns>the Group</returns>
        public HSSFShapeGroup CreateGroup(HSSFChildAnchor anchor)
        {
            HSSFShapeGroup group = new HSSFShapeGroup(this, anchor);
            group.Anchor = anchor;
            shapes.Add(group);
            return group;
        }

        /// <summary>
        /// Create a new simple shape Under this Group.
        /// </summary>
        /// <param name="anchor">the position of the shape.</param>
        /// <returns>the shape</returns>
        public HSSFSimpleShape CreateShape(HSSFChildAnchor anchor)
        {
            HSSFSimpleShape shape = new HSSFSimpleShape(this, anchor);
            shape.Anchor = anchor;
            shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Create a new textbox Under this Group.
        /// </summary>
        /// <param name="anchor">the position of the shape.</param>
        /// <returns>the textbox</returns>
        public HSSFTextbox CreateTextbox(HSSFChildAnchor anchor)
        {
            HSSFTextbox shape = new HSSFTextbox(this, anchor);
            shape.Anchor = anchor;
            shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Creates a polygon
        /// </summary>
        /// <param name="anchor">the client anchor describes how this Group Is attached
        /// to the sheet.</param>
        /// <returns>the newly Created shape.</returns>
        public HSSFPolygon CreatePolygon(HSSFChildAnchor anchor)
        {
            HSSFPolygon shape = new HSSFPolygon(this, anchor);
            shape.Anchor = anchor;
            shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Creates a picture.
        /// </summary>
        /// <param name="anchor">the client anchor describes how this Group Is attached
        /// to the sheet.</param>
        /// <param name="pictureIndex">Index of the picture.</param>
        /// <returns>the newly Created shape.</returns>
        public HSSFPicture CreatePicture(HSSFChildAnchor anchor, int pictureIndex)
        {
            HSSFPicture shape = new HSSFPicture(this, anchor);
            shape.Anchor = anchor;
            shape.PictureIndex=pictureIndex;
            shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Return all children contained by this shape.
        /// </summary>
        /// <value></value>
        public System.Collections.IList Children
        {
            get { return shapes; }
        }

        /// <summary>
        /// Sets the coordinate space of this Group.  All children are constrained
        /// to these coordinates.
        /// </summary>
        /// <param name="x1">The x1.</param>
        /// <param name="y1">The y1.</param>
        /// <param name="x2">The x2.</param>
        /// <param name="y2">The y2.</param>
        public void SetCoordinates(int x1, int y1, int x2, int y2)
        {
            this.x1 = x1;
            this.y1 = y1;
            this.x2 = x2;
            this.y2 = y2;
        }

        /// <summary>
        /// Gets The top left x coordinate of this Group.
        /// </summary>
        /// <value>The x1.</value>
        public int X1
        {
            get { return x1; }
        }

        /// <summary>
        /// Gets The top left y coordinate of this Group.
        /// </summary>
        /// <value>The y1.</value>
        public int Y1
        {
            get
            {
                return y1;
            }
        }

        /// <summary>
        /// Gets The bottom right x coordinate of this Group.
        /// </summary>
        /// <value>The x2.</value>
        public int X2
        {
            get
            {
                return x2;
            }
        }

        /// <summary>
        /// Gets the bottom right y coordinate of this Group.
        /// </summary>
        /// <value>The y2.</value>
        public int Y2
        {
            get
            {
                return y2;
            }
        }

        /// <summary>
        /// Count of all children and their childrens children.
        /// </summary>
        /// <value></value>
        public override int CountOfAllChildren
        {
            get
            {
                int count = shapes.Count;
                for (IEnumerator iterator = shapes.GetEnumerator(); iterator.MoveNext(); )
                {
                    HSSFShape shape = (HSSFShape)iterator.Current;
                    count += shape.CountOfAllChildren;
                }
                return count;
            }
        }

        internal override HSSFShape CloneShape()
        {
            throw new NotImplementedException("Use method cloneShape(HSSFPatriarch patriarch)");
        }

        internal HSSFShape CloneShape(HSSFPatriarch patriarch)
        {
            EscherContainerRecord spgrContainer = new EscherContainerRecord();
            spgrContainer.RecordId = (EscherContainerRecord.SPGR_CONTAINER);
            spgrContainer.Options = ((short)0x000F);
            EscherContainerRecord spContainer = new EscherContainerRecord();
            EscherContainerRecord cont = (EscherContainerRecord)GetEscherContainer().GetChildById(EscherContainerRecord.SP_CONTAINER);
            byte[] inSp = cont.Serialize();
            spContainer.FillFields(inSp, 0, new DefaultEscherRecordFactory());

            spgrContainer.AddChildRecord(spContainer);
            ObjRecord obj = null;
            if (null != this.GetObjRecord())
            {
                obj = (ObjRecord)this.GetObjRecord().CloneViaReserialise();
            }

            HSSFShapeGroup group = new HSSFShapeGroup(spgrContainer, obj);
            group.Patriarch = patriarch;

            foreach (HSSFShape shape in Children)
            {
                HSSFShape newShape;
                if (shape is HSSFShapeGroup)
                {
                    newShape = ((HSSFShapeGroup)shape).CloneShape(patriarch);
                }
                else
                {
                    newShape = shape.CloneShape();
                }
                group.AddShape(newShape);
                group.OnCreate(newShape);
            }
            return group;
        }
        private void OnCreate(HSSFShape shape)
        {
            if (this.Patriarch != null)
            {
                EscherContainerRecord spContainer = shape.GetEscherContainer();
                int shapeId = this.Patriarch.NewShapeId();
                shape.ShapeId = (shapeId);
                GetEscherContainer().AddChildRecord(spContainer);
                shape.AfterInsert(Patriarch);
                EscherSpRecord sp;
                if (shape is HSSFShapeGroup)
                {
                    sp = (EscherSpRecord)shape.GetEscherContainer().ChildContainers[0].GetChildById(EscherSpRecord.RECORD_ID);
                }
                else
                {
                    sp = (EscherSpRecord)shape.GetEscherContainer().GetChildById(EscherSpRecord.RECORD_ID);
                }
                sp.Flags = sp.Flags | EscherSpRecord.FLAG_CHILD;
            }
        }
        public void AddShape(HSSFShape shape)
        {
            shape.Patriarch = this.Patriarch;
            shape.Parent = this;
            shapes.Add(shape);
        }
        internal override void AfterInsert(HSSFPatriarch patriarch)
        {
            EscherAggregate agg = patriarch._getBoundAggregate();
            EscherContainerRecord containerRecord = (EscherContainerRecord)GetEscherContainer().GetChildById(EscherContainerRecord.SP_CONTAINER);
            agg.AssociateShapeToObjRecord(containerRecord.GetChildById(EscherClientDataRecord.RECORD_ID), GetObjRecord());
        }
        public override int ShapeId
        {
            get
            {
                EscherContainerRecord containerRecord = (EscherContainerRecord)GetEscherContainer().GetChildById(EscherContainerRecord.SP_CONTAINER);
                return ((EscherSpRecord)containerRecord.GetChildById(EscherSpRecord.RECORD_ID)).ShapeId;
            }
            set
            {
                EscherContainerRecord containerRecord = (EscherContainerRecord)GetEscherContainer().GetChildById(EscherContainerRecord.SP_CONTAINER);
                EscherSpRecord spRecord = (EscherSpRecord)containerRecord.GetChildById(EscherSpRecord.RECORD_ID);
                spRecord.ShapeId = value;
                CommonObjectDataSubRecord cod = (CommonObjectDataSubRecord)GetObjRecord().SubRecords[0];
                cod.ObjectId=(short)(value % 1024);
            }
        }
        protected override ObjRecord CreateObjRecord()
        {
            ObjRecord obj = new ObjRecord();
            CommonObjectDataSubRecord cmo = new CommonObjectDataSubRecord();
            cmo.ObjectType = (CommonObjectType.GROUP);
            cmo.IsLocked=(true);
            cmo.IsPrintable=(true);
            cmo.IsAutoFill=(true);
            cmo.IsAutoline=(true);
            GroupMarkerSubRecord gmo = new GroupMarkerSubRecord();
            EndSubRecord end = new EndSubRecord();
            obj.AddSubRecord(cmo);
            obj.AddSubRecord(gmo);
            obj.AddSubRecord(end);
            return obj;
        }

        protected override EscherContainerRecord CreateSpContainer()
        {
            EscherContainerRecord spgrContainer = new EscherContainerRecord();
            EscherContainerRecord spContainer = new EscherContainerRecord();
            EscherSpgrRecord spgr = new EscherSpgrRecord();
            EscherSpRecord sp = new EscherSpRecord();
            EscherOptRecord opt = new EscherOptRecord();
            EscherRecord anchor;
            EscherClientDataRecord clientData = new EscherClientDataRecord();

            spgrContainer.RecordId=(EscherContainerRecord.SPGR_CONTAINER);
            spgrContainer.Options=((short)0x000F);
            spContainer.RecordId=(EscherContainerRecord.SP_CONTAINER);
            spContainer.Options=(short)0x000F;
            spgr.RecordId=(EscherSpgrRecord.RECORD_ID);
            spgr.Options=(short)0x0001;
            spgr.RectX1=(0);
            spgr.RectY1=(0);
            spgr.RectX2=(1023);
            spgr.RectY2=(255);
            sp.RecordId=(EscherSpRecord.RECORD_ID);
            sp.Options=(short)0x0002;
            if (this.Anchor is HSSFClientAnchor)
            {
                sp.Flags = (EscherSpRecord.FLAG_GROUP | EscherSpRecord.FLAG_HAVEANCHOR);
            }
            else
            {
                sp.Flags = (EscherSpRecord.FLAG_GROUP | EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_CHILD);
            }
            opt.RecordId = (EscherOptRecord.RECORD_ID);
            opt.Options = ((short)0x0023);
            opt.AddEscherProperty(new EscherBoolProperty(EscherProperties.PROTECTION__LOCKAGAINSTGROUPING, 0x00040004));
            opt.AddEscherProperty(new EscherBoolProperty(EscherProperties.GROUPSHAPE__PRINT, 0x00080000));

            anchor = Anchor.GetEscherAnchor();
            clientData.RecordId = (EscherClientDataRecord.RECORD_ID);
            clientData.Options = ((short)0x0000);

            spgrContainer.AddChildRecord(spContainer);
            spContainer.AddChildRecord(spgr);
            spContainer.AddChildRecord(sp);
            spContainer.AddChildRecord(opt);
            spContainer.AddChildRecord(anchor);
            spContainer.AddChildRecord(clientData);
            return spgrContainer;
        }

        internal override void AfterRemove(HSSFPatriarch patriarch)
        {
            patriarch._getBoundAggregate().RemoveShapeToObjRecord(GetEscherContainer().ChildContainers[0]
                    .GetChildById(EscherClientDataRecord.RECORD_ID));
            for (int i = 0; i < shapes.Count; i++)
            {
                HSSFShape shape = (HSSFShape)shapes[i];
                RemoveShape(shape);
                shape.AfterRemove(Patriarch);
            }
            shapes.Clear();
        }
        public bool RemoveShape(HSSFShape shape)
        {
            bool isRemoved = GetEscherContainer().RemoveChildRecord(shape.GetEscherContainer());
            if (isRemoved)
            {
                shape.AfterRemove(this.Patriarch);
                shapes.Remove(shape);
            }
            return isRemoved;
        }
    }
}