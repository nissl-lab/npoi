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

using NPOI.ddf.*;
using NPOI.util.POILogger;



/**
 * Represents an AutoShape.
 * <p>
 * AutoShapes are Drawing objects with a particular shape that may be customized through smart resizing and adjustments.
 * See {@link ShapeTypes}
 * </p>
 *
 *  @author Yegor Kozlov
 */
public class AutoShape : TextShape {

    protected AutoShape(EscherContainerRecord escherRecord, Shape parent){
        base(escherRecord, parent);
    }

    public AutoShape(int type, Shape parent){
        base(null, parent);
        _escherContainer = CreateSpContainer(type, parent is ShapeGroup);
    }

    public AutoShape(int type){
        this(type, null);
    }

    protected EscherContainerRecord CreateSpContainer(int shapeType, bool IsChild){
        _escherContainer = super.CreateSpContainer(isChild);

        SetShapeType(shapeType);

        //set default properties for an autoshape
        SetEscherProperty(EscherProperties.PROTECTION__LOCKAGAINSTGROUPING, 0x40000);
        SetEscherProperty(EscherProperties.FILL__FILLCOLOR, 0x8000004);
        SetEscherProperty(EscherProperties.FILL__FILLCOLOR, 0x8000004);
        SetEscherProperty(EscherProperties.FILL__FILLBACKCOLOR, 0x8000000);
        SetEscherProperty(EscherProperties.FILL__NOFILLHITTEST, 0x100010);
        SetEscherProperty(EscherProperties.LINESTYLE__COLOR, 0x8000001);
        SetEscherProperty(EscherProperties.LINESTYLE__NOLINEDRAWDASH, 0x80008);
        SetEscherProperty(EscherProperties.SHADOWSTYLE__COLOR, 0x8000002);

        return _escherContainer;
    }

    protected void SetDefaultTextProperties(TextRun _txtRun){
        SetVerticalAlignment(TextBox.AnchorMiddle);
        SetHorizontalAlignment(TextBox.AlignCenter);
        SetWordWrap(TextBox.WrapNone);
    }

    /**
     * Gets adjust value which controls smart resizing of the auto-shape.
     *
     * <p>
     * The adjustment values are given in shape coordinates:
     * the origin is at the top-left, positive-x is to the right, positive-y is down.
     * The region from (0,0) to (S,S) maps to the geometry box of the shape (S=21600 is a constant).
     * </p>
     *
     * @param idx the adjust index in the [0, 9] range
     * @return the adjustment value
     */
    public int GetAdjustmentValue(int idx){
        if(idx < 0 || idx > 9) throw new ArgumentException("The index of an adjustment value must be in the [0, 9] range");

        return GetEscherProperty((short)(EscherProperties.GEOMETRY__ADJUSTVALUE + idx));
    }

    /**
     * Sets adjust value which controls smart resizing of the auto-shape.
     *
     * <p>
     * The adjustment values are given in shape coordinates:
     * the origin is at the top-left, positive-x is to the right, positive-y is down.
     * The region from (0,0) to (S,S) maps to the geometry box of the shape (S=21600 is a constant).
     * </p>
     *
     * @param idx the adjust index in the [0, 9] range
     * @param val the adjustment value
     */
    public void SetAdjustmentValue(int idx, int val){
        if(idx < 0 || idx > 9) throw new ArgumentException("The index of an adjustment value must be in the [0, 9] range");

        SetEscherProperty((short)(EscherProperties.GEOMETRY__ADJUSTVALUE + idx), val);
    }

    public java.awt.Shape GetOutline(){
        ShapeOutline outline = AutoShapes.GetShapeOutline(getShapeType());
        Rectangle2D anchor = GetLogicalAnchor2D();
        if(outline == null){
            logger.log(POILogger.WARN, "Outline not found for " + ShapeTypes.typeName(getShapeType()));
            return anchor;
        }
        java.awt.Shape shape = outline.GetOutline(this);
        return AutoShapes.transform(shape, anchor);
    }

}





