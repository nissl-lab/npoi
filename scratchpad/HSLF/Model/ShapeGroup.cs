/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HSLF.Model;








using NPOI.ddf.EscherChildAnchorRecord;
using NPOI.ddf.EscherClientAnchorRecord;
using NPOI.ddf.EscherContainerRecord;
using NPOI.ddf.EscherRecord;
using NPOI.ddf.EscherSpRecord;
using NPOI.ddf.EscherSpgrRecord;
using NPOI.util.LittleEndian;
using NPOI.util.POILogger;

/**
 *  Represents a group of shapes.
 *
 * @author Yegor Kozlov
 */
public class ShapeGroup : Shape{

    /**
      * Create a new ShapeGroup. This constructor is used when a new shape is Created.
      *
      */
    public ShapeGroup(){
        this(null, null);
        _escherContainer = CreateSpContainer(false);
    }

    /**
      * Create a ShapeGroup object and Initilize it from the supplied Record Container.
      *
      * @param escherRecord       <code>EscherSpContainer</code> Container which holds information about this shape
      * @param parent    the parent of the shape
      */
    protected ShapeGroup(EscherContainerRecord escherRecord, Shape parent){
        base(escherRecord, parent);
    }

    /**
     * @return the shapes Contained in this group Container
     */
    public Shape[] GetShapes() {
    	// Out escher Container record should contain several
        //  SpContainers, the first of which is the group shape itself
        Iterator<EscherRecord> iter = _escherContainer.GetChildIterator();

        // Don't include the first SpContainer, it is always NotPrimitive
        if (iter.HasNext()) {
        	iter.next();
        }
        List<Shape> shapeList = new List<Shape>();
        while (iter.HasNext()) {
        	EscherRecord r = iter.next();
        	if(r is EscherContainerRecord) {
        		// Create the Shape for it
        		EscherContainerRecord Container = (EscherContainerRecord)r;
        		Shape shape = ShapeFactory.CreateShape(Container, this);
                shape.SetSheet(getSheet());
        		shapeList.Add( shape );
        	} else {
        		// Should we do anything special with these non
        		//  Container records?
        		logger.log(POILogger.ERROR, "Shape Contained non Container escher record, was " + r.GetType().Name);
        	}
        }

        // Put the shapes into an array, and return
        Shape[] shapes = shapeList.ToArray(new Shape[shapeList.Count]);
        return shapes;
    }

    /**
     * Sets the anchor (the bounding box rectangle) of this shape.
     * All coordinates should be expressed in Master units (576 dpi).
     *
     * @param anchor new anchor
     */
    public void SetAnchor(java.awt.Rectangle anchor){

        EscherContainerRecord spContainer = (EscherContainerRecord)_escherContainer.GetChild(0);

        EscherClientAnchorRecord clientAnchor = (EscherClientAnchorRecord)getEscherChild(spContainer, EscherClientAnchorRecord.RECORD_ID);
        //hack. internal variable EscherClientAnchorRecord.shortRecord can be
        //Initialized only in FillFields(). We need to Set shortRecord=false;
        byte[] header = new byte[16];
        LittleEndian.PutUshort(header, 0, 0);
        LittleEndian.PutUshort(header, 2, 0);
        LittleEndian.PutInt(header, 4, 8);
        clientAnchor.FillFields(header, 0, null);

        clientAnchor.SetFlag((short)(anchor.y*MASTER_DPI/POINT_DPI));
        clientAnchor.SetCol1((short)(anchor.x*MASTER_DPI/POINT_DPI));
        clientAnchor.SetDx1((short)((anchor.width + anchor.x)*MASTER_DPI/POINT_DPI));
        clientAnchor.SetRow1((short)((anchor.height + anchor.y)*MASTER_DPI/POINT_DPI));

        EscherSpgrRecord spgr = (EscherSpgrRecord)getEscherChild(spContainer, EscherSpgrRecord.RECORD_ID);

        spgr.SetRectX1(anchor.x*MASTER_DPI/POINT_DPI);
        spgr.SetRectY1(anchor.y*MASTER_DPI/POINT_DPI);
        spgr.SetRectX2((anchor.x + anchor.width)*MASTER_DPI/POINT_DPI);
        spgr.SetRectY2((anchor.y + anchor.height)*MASTER_DPI/POINT_DPI);
    }

    /**
     * Sets the coordinate space of this group.  All children are constrained
     * to these coordinates.
     *
     * @param anchor the coordinate space of this group
     */
    public void SetCoordinates(Rectangle2D anchor){
        EscherContainerRecord spContainer = (EscherContainerRecord)_escherContainer.GetChild(0);
        EscherSpgrRecord spgr = (EscherSpgrRecord)getEscherChild(spContainer, EscherSpgrRecord.RECORD_ID);

        int x1 = (int)Math.round(anchor.GetX()*MASTER_DPI/POINT_DPI);
        int y1 = (int)Math.round(anchor.GetY()*MASTER_DPI/POINT_DPI);
        int x2 = (int)Math.round((anchor.GetX() + anchor.Width)*MASTER_DPI/POINT_DPI);
        int y2 = (int)Math.round((anchor.GetY() + anchor.Height)*MASTER_DPI/POINT_DPI);

        spgr.SetRectX1(x1);
        spgr.SetRectY1(y1);
        spgr.SetRectX2(x2);
        spgr.SetRectY2(y2);

    }

    /**
     * Gets the coordinate space of this group.  All children are constrained
     * to these coordinates.
     *
     * @return the coordinate space of this group
     */
    public Rectangle2D GetCoordinates(){
        EscherContainerRecord spContainer = (EscherContainerRecord)_escherContainer.GetChild(0);
        EscherSpgrRecord spgr = (EscherSpgrRecord)getEscherChild(spContainer, EscherSpgrRecord.RECORD_ID);

        Rectangle2D.Float anchor = new Rectangle2D.Float();
        anchor.x = (float)spgr.GetRectX1()*POINT_DPI/MASTER_DPI;
        anchor.y = (float)spgr.GetRectY1()*POINT_DPI/MASTER_DPI;
        anchor.width = (float)(spgr.GetRectX2() - spgr.GetRectX1())*POINT_DPI/MASTER_DPI;
        anchor.height = (float)(spgr.GetRectY2() - spgr.GetRectY1())*POINT_DPI/MASTER_DPI;

        return anchor;
    }

    /**
     * Create a new ShapeGroup and create an instance of <code>EscherSpgrContainer</code> which represents a group of shapes
     */
    protected EscherContainerRecord CreateSpContainer(bool IsChild) {
        EscherContainerRecord spgr = new EscherContainerRecord();
        spgr.SetRecordId(EscherContainerRecord.SPGR_CONTAINER);
        spgr.SetOptions((short)15);

        //The group itself is a shape, and always appears as the first EscherSpContainer in the group Container.
        EscherContainerRecord spcont = new EscherContainerRecord();
        spcont.SetRecordId(EscherContainerRecord.SP_CONTAINER);
        spcont.SetOptions((short)15);

        EscherSpgrRecord spg = new EscherSpgrRecord();
        spg.SetOptions((short)1);
        spcont.AddChildRecord(spg);

        EscherSpRecord sp = new EscherSpRecord();
        short type = (ShapeTypes.NotPrimitive << 4) + 2;
        sp.SetOptions(type);
        sp.SetFlags(EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_GROUP);
        spcont.AddChildRecord(sp);

        EscherClientAnchorRecord anchor = new EscherClientAnchorRecord();
        spcont.AddChildRecord(anchor);

        spgr.AddChildRecord(spcont);
        return spgr;
    }

    /**
     * Add a shape to this group.
     *
     * @param shape - the Shape to add
     */
    public void AddShape(Shape shape){
        _escherContainer.AddChildRecord(shape.GetSpContainer());

        Sheet sheet = Sheet;
        shape.SetSheet(sheet);
        shape.SetShapeId(sheet.allocateShapeId());
        shape.afterInsert(sheet);
    }

    /**
     * Moves this <code>ShapeGroup</code> to the specified location.
     * <p>
     * @param x the x coordinate of the top left corner of the shape in new location
     * @param y the y coordinate of the top left corner of the shape in new location
     */
    public void moveTo(int x, int y){
        java.awt.Rectangle anchor = GetAnchor();
        int dx = x - anchor.x;
        int dy = y - anchor.y;
        anchor.translate(dx, dy);
        SetAnchor(anchor);

        Shape[] shape = GetShapes();
        for (int i = 0; i < shape.Length; i++) {
            java.awt.Rectangle chanchor = shape[i].GetAnchor();
            chanchor.translate(dx, dy);
            shape[i].SetAnchor(chanchor);
        }
    }

    /**
     * Returns the anchor (the bounding box rectangle) of this shape group.
     * All coordinates are expressed in points (72 dpi).
     *
     * @return the anchor of this shape group
     */
    public Rectangle2D GetAnchor2D(){
        EscherContainerRecord spContainer = (EscherContainerRecord)_escherContainer.GetChild(0);
        EscherClientAnchorRecord clientAnchor = (EscherClientAnchorRecord)getEscherChild(spContainer, EscherClientAnchorRecord.RECORD_ID);
        Rectangle2D.Float anchor = new Rectangle2D.Float();
        if(clientAnchor == null){
            logger.log(POILogger.INFO, "EscherClientAnchorRecord was not found for shape group. Searching for EscherChildAnchorRecord.");
            EscherChildAnchorRecord rec = (EscherChildAnchorRecord)getEscherChild(spContainer, EscherChildAnchorRecord.RECORD_ID);
            anchor = new Rectangle2D.Float(
                (float)rec.GetDx1()*POINT_DPI/MASTER_DPI,
                (float)rec.GetDy1()*POINT_DPI/MASTER_DPI,
                (float)(rec.GetDx2()-rec.GetDx1())*POINT_DPI/MASTER_DPI,
                (float)(rec.GetDy2()-rec.GetDy1())*POINT_DPI/MASTER_DPI
            );
        } else {
            anchor.x = (float)clientAnchor.GetCol1()*POINT_DPI/MASTER_DPI;
            anchor.y = (float)clientAnchor.GetFlag()*POINT_DPI/MASTER_DPI;
            anchor.width = (float)(clientAnchor.GetDx1() - clientAnchor.GetCol1())*POINT_DPI/MASTER_DPI ;
            anchor.height = (float)(clientAnchor.GetRow1() - clientAnchor.GetFlag())*POINT_DPI/MASTER_DPI;
        }

        return anchor;
    }

    /**
     * Return type of the shape.
     * In most cases shape group type is {@link NPOI.HSLF.Model.ShapeTypes#NotPrimitive}
     *
     * @return type of the shape.
     */
    public int GetShapeType(){
        EscherContainerRecord groupInfoContainer = (EscherContainerRecord)_escherContainer.GetChild(0);
        EscherSpRecord spRecord = groupInfoContainer.GetChildById(EscherSpRecord.RECORD_ID);
        return spRecord.GetOptions() >> 4;
    }

    /**
     * Returns <code>null</code> - shape groups can't have hyperlinks
     *
     * @return <code>null</code>.
     */
     public Hyperlink GetHyperlink(){
        return null;
    }

    public void Draw(Graphics2D graphics){

        AffineTransform at = graphics.GetTransform();

        Shape[] sh = GetShapes();
        for (int i = 0; i < sh.Length; i++) {
            sh[i].Draw(graphics);
        }

        graphics.SetTransform(at);
    }
}





