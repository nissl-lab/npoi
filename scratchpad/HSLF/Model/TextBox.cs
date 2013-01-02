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

/**
 * Represents a TextFrame shape in PowerPoint.
 * <p>
 * Contains the text in a text frame as well as the properties and methods
 * that control alignment and anchoring of the text.
 * </p>
 *
 * @author Yegor Kozlov
 */
public class TextBox : TextShape {

    /**
     * Create a TextBox object and Initialize it from the supplied Record Container.
     *
     * @param escherRecord       <code>EscherSpContainer</code> Container which holds information about this shape
     * @param parent    the parent of the shape
     */
   protected TextBox(EscherContainerRecord escherRecord, Shape parent){
        base(escherRecord, parent);

    }

    /**
     * Create a new TextBox. This constructor is used when a new shape is Created.
     *
     * @param parent    the parent of this Shape. For example, if this text box is a cell
     * in a table then the parent is Table.
     */
    public TextBox(Shape parent){
        base(parent);
    }

    /**
     * Create a new TextBox. This constructor is used when a new shape is Created.
     *
     */
    public TextBox(){
        this(null);
    }

    /**
     * Create a new TextBox and Initialize its internal structures
     *
     * @return the Created <code>EscherContainerRecord</code> which holds shape data
     */
    protected EscherContainerRecord CreateSpContainer(bool IsChild){
        _escherContainer = super.CreateSpContainer(isChild);

        SetShapeType(ShapeTypes.TextBox);

        //set default properties for a TextBox
        SetEscherProperty(EscherProperties.FILL__FILLCOLOR, 0x8000004);
        SetEscherProperty(EscherProperties.FILL__FILLBACKCOLOR, 0x8000000);
        SetEscherProperty(EscherProperties.FILL__NOFILLHITTEST, 0x100000);
        SetEscherProperty(EscherProperties.LINESTYLE__COLOR, 0x8000001);
        SetEscherProperty(EscherProperties.LINESTYLE__NOLINEDRAWDASH, 0x80000);
        SetEscherProperty(EscherProperties.SHADOWSTYLE__COLOR, 0x8000002);

        _txtrun = CreateTextRun();

        return _escherContainer;
    }

    protected void SetDefaultTextProperties(TextRun _txtRun){
        SetVerticalAlignment(TextBox.AnchorTop);
        SetEscherProperty(EscherProperties.TEXT__SIZE_TEXT_TO_FIT_SHAPE, 0x20002);
    }

}





