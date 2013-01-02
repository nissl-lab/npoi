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
using NPOI.HSLF.record.OEPlaceholderAtom;
using NPOI.HSLF.exceptions.HSLFException;



/**
 * Represents a Placeholder in PowerPoint.
 *
 * @author Yegor Kozlov
 */
public class Placeholder : TextBox {

    protected Placeholder(EscherContainerRecord escherRecord, Shape parent){
        base(escherRecord, parent);
    }

    public Placeholder(Shape parent){
        base(parent);
    }

    public Placeholder(){
        base();
    }

    /**
     * Create a new Placeholder and Initialize internal structures
     *
     * @return the Created <code>EscherContainerRecord</code> which holds shape data
     */
    protected EscherContainerRecord CreateSpContainer(bool IsChild){
        _escherContainer = super.CreateSpContainer(isChild);

        EscherSpRecord spRecord = _escherContainer.GetChildById(EscherSpRecord.RECORD_ID);
        spRecord.SetFlags(EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HAVEMASTER);

        EscherClientDataRecord cldata = new EscherClientDataRecord();
        cldata.SetOptions((short)15);

        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);

        //Placeholders can't be grouped
        SetEscherProperty(opt, EscherProperties.PROTECTION__LOCKAGAINSTGROUPING, 262144);

        //OEPlaceholderAtom tells powerpoint that this shape is a placeholder
        //
        OEPlaceholderAtom oep = new OEPlaceholderAtom();
        /**
         * Extarct from MSDN:
         *
         * There is a special case when the placeholder does not have a position in the layout.
         * This occurs when the user has moved the placeholder from its original position.
         * In this case the placeholder ID is -1.
         */
        oep.SetPlacementId(-1);

        oep.SetPlaceholderId(OEPlaceholderAtom.Body);

        //convert hslf into ddf record
        MemoryStream out = new MemoryStream();
        try {
            oep.WriteOut(out);
        } catch(Exception e){
            throw new HSLFException(e);
        }
        cldata.SetRemainingData(out.ToArray());

        //append placeholder Container before EscherTextboxRecord
        _escherContainer.AddChildBefore(cldata, EscherTextboxRecord.RECORD_ID);

        return _escherContainer;
    }
}





