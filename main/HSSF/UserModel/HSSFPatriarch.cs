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
    using System;
    using System.Collections;
    using NPOI.DDF;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    using NPOI.SS.UserModel;
    using System.Collections.Generic;
    using NPOI.HSSF.Model;
    using NPOI.SS.Util;
    using NPOI.POIFS.FileSystem;
    using System.IO;

    /// <summary>
    /// The patriarch is the toplevel container for shapes in a sheet.  It does
    /// little other than act as a container for other shapes and Groups.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class HSSFPatriarch : HSSFShapeContainer, IDrawing
    {
        //private static POILogger log = POILogFactory.GetLogger(typeof(HSSFPatriarch));
        List<HSSFShape> _shapes = new List<HSSFShape>();
        private HSSFSheet _sheet;
        private EscherSpgrRecord _spgrRecord;
        private EscherContainerRecord _mainSpgrContainer;

        /**
         * The EscherAggregate we have been bound to.
         * (This will handle writing us out into records,
         *  and building up our shapes from the records)
         */
        private EscherAggregate _boundAggregate;

        /// <summary>
        /// Creates the patriarch.
        /// </summary>
        /// <param name="sheet">the sheet this patriarch is stored in.</param>
        /// <param name="boundAggregate">The bound aggregate.</param>
        public HSSFPatriarch(HSSFSheet sheet, EscherAggregate boundAggregate)
        {
            _boundAggregate = boundAggregate;
            _sheet = sheet;
            _mainSpgrContainer = _boundAggregate.GetEscherContainer().ChildContainers[0];
            EscherContainerRecord spContainer = (EscherContainerRecord)_boundAggregate.GetEscherContainer()
                    .ChildContainers[0].GetChild(0);
            _spgrRecord = (EscherSpgrRecord)spContainer.GetChildById(EscherSpgrRecord.RECORD_ID);
            BuildShapeTree();
        }


        public static HSSFPatriarch CreatePatriarch(HSSFPatriarch patriarch, HSSFSheet sheet)
        {
            HSSFPatriarch newPatriarch = new HSSFPatriarch(sheet, new EscherAggregate(true));
            newPatriarch.AfterCreate();
            foreach (HSSFShape shape in patriarch.Children)
            {
                HSSFShape newShape;
                if (shape is HSSFShapeGroup)
                {
                    newShape = ((HSSFShapeGroup)shape).CloneShape(newPatriarch);
                }
                else
                {
                    newShape = shape.CloneShape();
                }
                newPatriarch.OnCreate(newShape);
                newPatriarch.AddShape(newShape);
            }
            return newPatriarch;
        }

        /**
     * check if any shapes contain wrong data
     * At now(13.08.2010) check if patriarch contains 2 or more comments with same coordinates
     */
        protected internal void PreSerialize()
        {
            Dictionary<int, NoteRecord> tailRecords = _boundAggregate.TailRecords;
            /*
             * contains coordinates of comments we iterate over
             */
            Hashtable coordinates = new Hashtable(tailRecords.Count);
            foreach (NoteRecord rec in tailRecords.Values)
            {
                String noteRef = new CellReference(rec.Row, rec.Column).FormatAsString(); // A1-style notation
                if (coordinates.Contains(noteRef))
                {
                    throw new InvalidOperationException("found multiple cell comments for cell " + noteRef);
                }
                else
                {
                    coordinates.Add(noteRef, null);
                }
            }
        }

        /**
         * @param shape to be removed
         * @return true of shape is removed
         */
        public bool RemoveShape(HSSFShape shape)
        {
            bool isRemoved = _mainSpgrContainer.RemoveChildRecord(shape.GetEscherContainer());
            if (isRemoved)
            {
                shape.AfterRemove(this);
                _shapes.Remove(shape);
            }
            return isRemoved;
        }

        internal void AfterCreate()
        {
            DrawingManager2 drawingManager = ((HSSFWorkbook)_sheet.Workbook).Workbook.DrawingManager;
            short dgId = drawingManager.FindNewDrawingGroupId();
            _boundAggregate.SetDgId(dgId);
            _boundAggregate.SetMainSpRecordId(NewShapeId());
            drawingManager.IncrementDrawingsSaved();
        }

        /// <summary>
        /// Creates a new Group record stored Under this patriarch.
        /// </summary>
        /// <param name="anchor">the client anchor describes how this Group is attached
        /// to the sheet.</param>
        /// <returns>the newly created Group.</returns>
        public HSSFShapeGroup CreateGroup(HSSFClientAnchor anchor)
        {
            HSSFShapeGroup group = new HSSFShapeGroup(null, anchor);

            AddShape(group);
            OnCreate(group);
            return group;
        }

        /// <summary>
        /// Creates a simple shape.  This includes such shapes as lines, rectangles,
        /// and ovals.
        /// Note: Microsoft Excel seems to sometimes disallow 
        /// higher y1 than y2 or higher x1 than x2 in the anchor, you might need to 
        /// reverse them and draw shapes vertically or horizontally flipped! 
        /// </summary>
        /// <param name="anchor">the client anchor describes how this Group is attached
        /// to the sheet.</param>
        /// <returns>the newly created shape.</returns>
        public HSSFSimpleShape CreateSimpleShape(HSSFClientAnchor anchor)
        {
            HSSFSimpleShape shape = new HSSFSimpleShape(null, anchor);

            AddShape(shape);
            //open existing file
            OnCreate(shape);
            return shape;
        }

        /// <summary>
        /// Creates a picture.
        /// </summary>
        /// <param name="anchor">the client anchor describes how this Group is attached
        /// to the sheet.</param>
        /// <param name="pictureIndex">Index of the picture.</param>
        /// <returns>the newly created shape.</returns>
        public IPicture CreatePicture(HSSFClientAnchor anchor, int pictureIndex)
        {
            HSSFPicture shape = new HSSFPicture(null, (HSSFClientAnchor)anchor);
            shape.PictureIndex = pictureIndex;
            AddShape(shape);

            //open existing file
            OnCreate(shape);
            return shape;
        }

        /// <summary>
        /// CreatePicture
        /// </summary>
        /// <param name="anchor">the client anchor describes how this picture is attached to the sheet.</param>
        /// <param name="pictureIndex">the index of the picture in the workbook collection of pictures.</param>
        /// <returns>return newly created shape</returns>
        public IPicture CreatePicture(IClientAnchor anchor, int pictureIndex)
        {
            return CreatePicture((HSSFClientAnchor)anchor, pictureIndex);
        }

        /**
     * Adds a new OLE Package Shape 
     * 
     * @param anchor       the client anchor describes how this picture is
     *                     attached to the sheet.
     * @param storageId    the storageId returned by {@Link HSSFWorkbook.AddOlePackage}
     * @param pictureIndex the index of the picture (used as preview image) in the
     *                     workbook collection of pictures.
     *
     * @return newly Created shape
     */
        public HSSFObjectData CreateObjectData(HSSFClientAnchor anchor, int storageId, int pictureIndex)
        {
            ObjRecord obj = new ObjRecord();

            CommonObjectDataSubRecord ftCmo = new CommonObjectDataSubRecord();
            ftCmo.ObjectType = (/*setter*/CommonObjectType.Picture);
            // ftCmo.ObjectId=(/*setter*/oleShape.ShapeId); ... will be Set by onCreate(...)
            ftCmo.IsLocked = (/*setter*/true);
            ftCmo.IsPrintable = (/*setter*/true);
            ftCmo.IsAutoFill = (/*setter*/true);
            ftCmo.IsAutoline = (/*setter*/true);
            ftCmo.Reserved1 = (/*setter*/0);
            ftCmo.Reserved2 = (/*setter*/0);
            ftCmo.Reserved3 = (/*setter*/0);
            obj.AddSubRecord(ftCmo);

            // FtCf (pictFormat) 
            FtCfSubRecord ftCf = new FtCfSubRecord();
            HSSFPictureData pictData = Sheet.Workbook.GetAllPictures()[(pictureIndex - 1)] as HSSFPictureData;
            switch ((PictureType)pictData.Format)
            {
                case PictureType.WMF:
                case PictureType.EMF:
                    // this needs patch #49658 to be applied to actually work 
                    ftCf.Flags = (/*setter*/FtCfSubRecord.METAFILE_BIT);
                    break;
                case PictureType.DIB:
                case PictureType.PNG:
                case PictureType.JPEG:
                case PictureType.PICT:
                    ftCf.Flags = (/*setter*/FtCfSubRecord.BITMAP_BIT);
                    break;
            }
            obj.AddSubRecord(ftCf);
            // FtPioGrbit (pictFlags)
            FtPioGrbitSubRecord ftPioGrbit = new FtPioGrbitSubRecord();
            ftPioGrbit.SetFlagByBit(FtPioGrbitSubRecord.AUTO_PICT_BIT, true);
            obj.AddSubRecord(ftPioGrbit);

            EmbeddedObjectRefSubRecord ftPictFmla = new EmbeddedObjectRefSubRecord();
            ftPictFmla.SetUnknownFormulaData(new byte[] { 2, 0, 0, 0, 0 });
            ftPictFmla.OLEClassName = (/*setter*/"Paket");
            ftPictFmla.SetStorageId(storageId);

            obj.AddSubRecord(ftPictFmla);
            obj.AddSubRecord(new EndSubRecord());

            String entryName = "MBD" + HexDump.ToHex(storageId);
            DirectoryEntry oleRoot;
            try
            {
                DirectoryNode dn = (_sheet.Workbook as HSSFWorkbook).RootDirectory;
                if (dn == null) throw new FileNotFoundException();
                oleRoot = (DirectoryEntry)dn.GetEntry(entryName);
            }
            catch (FileNotFoundException e)
            {
                throw new InvalidOperationException("trying to add ole shape without actually Adding data first - use HSSFWorkbook.AddOlePackage first", e);
            }

            // create picture shape, which need to be minimal modified for oleshapes
            HSSFPicture shape = new HSSFPicture(null, anchor);
            shape.PictureIndex = (/*setter*/pictureIndex);
            EscherContainerRecord spContainer = shape.GetEscherContainer();
            EscherSpRecord spRecord = spContainer.GetChildById(EscherSpRecord.RECORD_ID) as EscherSpRecord;
            spRecord.Flags = (/*setter*/spRecord.Flags | EscherSpRecord.FLAG_OLESHAPE);

            HSSFObjectData oleShape = new HSSFObjectData(spContainer, obj, oleRoot);
            AddShape(oleShape);
            OnCreate(oleShape);


            return oleShape;
        }

        /// <summary>
        /// Creates a polygon
        /// </summary>
        /// <param name="anchor">the client anchor describes how this Group is attached
        /// to the sheet.</param>
        /// <returns>the newly Created shape.</returns>
        public HSSFPolygon CreatePolygon(IClientAnchor anchor)
        {
            HSSFPolygon shape = new HSSFPolygon(null, (HSSFAnchor)anchor);
            AddShape(shape);
            OnCreate(shape);
            return shape;
        }

        /// <summary>
        /// Constructs a textbox Under the patriarch.
        /// </summary>
        /// <param name="anchor">the client anchor describes how this Group is attached
        /// to the sheet.</param>
        /// <returns>the newly Created textbox.</returns>
        public HSSFSimpleShape CreateTextbox(IClientAnchor anchor)
        {
            HSSFTextbox shape = new HSSFTextbox(null, (HSSFAnchor)anchor);
            AddShape(shape);
            OnCreate(shape);
            return shape;
        }
        /**
         * Constructs a cell comment.
         *
         * @param anchor    the client anchor describes how this comment is attached
         *                  to the sheet.
         * @return      the newly created comment.
         */
        public HSSFComment CreateComment(HSSFAnchor anchor)
        {
            HSSFComment shape = new HSSFComment(null, anchor);
            AddShape(shape);
            OnCreate(shape);
            return shape;
        }
        /**
         * YK: used to create autofilters
         *
         * @see org.apache.poi.hssf.usermodel.HSSFSheet#setAutoFilter(int, int, int, int)
         */
        public HSSFSimpleShape CreateComboBox(HSSFAnchor anchor)
        {
            HSSFCombobox shape = new HSSFCombobox(null, anchor);
            AddShape(shape);
            OnCreate(shape);
            return shape;
        }

        /// <summary>
        /// Constructs a cell comment.
        /// </summary>
        /// <param name="anchor">the client anchor describes how this comment is attached
        /// to the sheet.</param>
        /// <returns>the newly created comment.</returns>
        public IComment CreateCellComment(IClientAnchor anchor)
        {
            return CreateComment((HSSFAnchor)anchor);
        }




        private void SetFlipFlags(HSSFShape shape)
        {
            EscherSpRecord sp = (EscherSpRecord)shape.GetEscherContainer().GetChildById(EscherSpRecord.RECORD_ID);
            if (shape.Anchor.IsHorizontallyFlipped)
            {
                sp.Flags = (sp.Flags | EscherSpRecord.FLAG_FLIPHORIZ);
            }
            if (shape.Anchor.IsVerticallyFlipped)
            {
                sp.Flags = (sp.Flags | EscherSpRecord.FLAG_FLIPVERT);
            }
        }
        /// <summary>
        /// Returns a list of all shapes contained by the patriarch.
        /// </summary>
        /// <value>The children.</value>
        public IList<HSSFShape> Children
        {
            get { return _shapes; }
        }

        /**
         * add a shape to this drawing
         */
        public void AddShape(HSSFShape shape)
        {
            shape.Patriarch = this;
            _shapes.Add(shape);
        }

        private void OnCreate(HSSFShape shape)
        {
            EscherContainerRecord spgrContainer =
                    _boundAggregate.GetEscherContainer().ChildContainers[0];

            EscherContainerRecord spContainer = shape.GetEscherContainer();
            int shapeId = NewShapeId();
            shape.ShapeId = shapeId;

            spgrContainer.AddChildRecord(spContainer);
            shape.AfterInsert(this);
            SetFlipFlags(shape);
        }
        /// <summary>
        /// Total count of all children and their children's children.
        /// </summary>
        /// <value>The count of all children.</value>
        public int CountOfAllChildren
        {
            get
            {
                int count = _shapes.Count;
                for (IEnumerator iterator = _shapes.GetEnumerator(); iterator.MoveNext(); )
                {
                    HSSFShape shape = (HSSFShape)iterator.Current;
                    count += shape.CountOfAllChildren;
                }
                return count;
            }
        }
        /// <summary>
        /// Sets the coordinate space of this Group.  All children are contrained
        /// to these coordinates.
        /// </summary>
        /// <param name="x1">The x1.</param>
        /// <param name="y1">The y1.</param>
        /// <param name="x2">The x2.</param>
        /// <param name="y2">The y2.</param>
        public void SetCoordinates(int x1, int y1, int x2, int y2)
        {
            _spgrRecord.RectY1 = (y1);
            _spgrRecord.RectY2 = (y2);
            _spgrRecord.RectX1 = (x1);
            _spgrRecord.RectX2 = (x2);
        }

        public void Clear()
        {
            List<HSSFShape> copy = new List<HSSFShape>(_shapes);
            foreach (HSSFShape shape in copy)
            {
                RemoveShape(shape);
            }
        }

        internal int NewShapeId()
        {
            DrawingManager2 dm = ((HSSFWorkbook)_sheet.Workbook).Workbook.DrawingManager;
            EscherDgRecord dg =
                   (EscherDgRecord)_boundAggregate.GetEscherContainer().GetChildById(EscherDgRecord.RECORD_ID);
            short drawingGroupId = dg.DrawingGroupId;
            return dm.AllocateShapeId(drawingGroupId, dg);
        }
        /// <summary>
        /// Does this HSSFPatriarch contain a chart?
        /// (Technically a reference to a chart, since they
        /// Get stored in a different block of records)
        /// FIXME - detect chart in all cases (only seems
        /// to work on some charts so far)
        /// </summary>
        /// <returns>
        /// 	<c>true</c> if this instance contains chart; otherwise, <c>false</c>.
        /// </returns>
        public bool ContainsChart()
        {
            // TODO - support charts properly in usermodel

            // We're looking for a EscherOptRecord
            EscherOptRecord optRecord = (EscherOptRecord)
                _boundAggregate.FindFirstWithId(EscherOptRecord.RECORD_ID);
            if (optRecord == null)
            {
                // No opt record, can't have chart
                return false;
            }

            for (IEnumerator it = optRecord.EscherProperties.GetEnumerator(); it.MoveNext(); )
            {
                EscherProperty prop = (EscherProperty)it.Current;
                if (prop.PropertyNumber == 896 && prop.IsComplex)
                {
                    EscherComplexProperty cp = (EscherComplexProperty)prop;
                    String str = StringUtil.GetFromUnicodeLE(cp.ComplexData);
                    //Console.Error.WriteLine(str);
                    if (str.Equals("Chart 1\0"))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// The top left x coordinate of this Group.
        /// </summary>
        /// <value>The x1.</value>
        public int X1
        {
            get { return _spgrRecord.RectX1; }
        }

        /// <summary>
        /// The top left y coordinate of this Group.
        /// </summary>
        /// <value>The y1.</value>
        public int Y1
        {
            get { return _spgrRecord.RectY1; }
        }

        /// <summary>
        /// The bottom right x coordinate of this Group.
        /// </summary>
        /// <value>The x2.</value>
        public int X2
        {
            get { return _spgrRecord.RectX2; }
        }

        /// <summary>
        /// The bottom right y coordinate of this Group.
        /// </summary>
        /// <value>The y2.</value>
        public int Y2
        {
            get { return _spgrRecord.RectY2; }
        }

        /// <summary>
        /// Returns the aggregate escher record we're bound to
        /// </summary>
        /// <returns></returns>
        internal EscherAggregate GetBoundAggregate()
        {
            return _boundAggregate;
        }
        /**
         * Creates a new client anchor and sets the top-left and bottom-right
         * coordinates of the anchor.
         *
         * @param dx1  the x coordinate in EMU within the first cell.
         * @param dy1  the y coordinate in EMU within the first cell.
         * @param dx2  the x coordinate in EMU within the second cell.
         * @param dy2  the y coordinate in EMU within the second cell.
         * @param col1 the column (0 based) of the first cell.
         * @param row1 the row (0 based) of the first cell.
         * @param col2 the column (0 based) of the second cell.
         * @param row2 the row (0 based) of the second cell.
         * @return the newly created client anchor
         */
        public IClientAnchor CreateAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2)
        {
            return new HSSFClientAnchor(dx1, dy1, dx2, dy2, (short)col1, row1, (short)col2, row2);
        }

        public IChart CreateChart(IClientAnchor anchor)
        {
            throw new RuntimeException("NotImplemented");
        }
        /**
     * create shape tree from existing escher records tree
     */
        public void BuildShapeTree()
        {
            EscherContainerRecord dgContainer = _boundAggregate.GetEscherContainer();
            if (dgContainer == null)
            {
                return;
            }
            EscherContainerRecord spgrConrainer = dgContainer.ChildContainers[0];
            IList<EscherContainerRecord> spgrChildren = spgrConrainer.ChildContainers;

            for (int i = 0; i < spgrChildren.Count; i++)
            {
                EscherContainerRecord spContainer = spgrChildren[i];
                if (i != 0)
                {
                    HSSFShapeFactory.CreateShapeTree(spContainer, _boundAggregate, this, ((HSSFWorkbook)_sheet.Workbook).RootDirectory);
                }
            }
        }

        public List<HSSFShape> GetShapes()
        {
            return _shapes;
        }
        public IEnumerator<HSSFShape> GetEnumerator()
        {
            return _shapes.GetEnumerator();
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _shapes.GetEnumerator();
        }
        protected internal HSSFSheet Sheet
        {
            get
            {
                return _sheet;
            }
        }
    }
}