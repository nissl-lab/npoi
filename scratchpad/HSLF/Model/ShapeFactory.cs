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
using NPOI.util.POILogFactory;
using NPOI.HSLF.record.*;




/**
 * Create a <code>Shape</code> object depending on its type
 *
 * @author Yegor Kozlov
 */
public class ShapeFactory {
    // For logging
    protected static POILogger logger = POILogFactory.GetLogger(ShapeFactory.class);

    /**
     * Create a new shape from the data provided.
     */
    public static Shape CreateShape(EscherContainerRecord spContainer, Shape parent){
        if (spContainer.GetRecordId() == EscherContainerRecord.SPGR_CONTAINER){
            return CreateShapeGroup(spContainer, parent);
        }
        return CreateSimpeShape(spContainer, parent);
    }

    public static ShapeGroup CreateShapeGroup(EscherContainerRecord spContainer, Shape parent){
        ShapeGroup group = null;
        EscherRecord opt = Shape.GetEscherChild((EscherContainerRecord)spContainer.GetChild(0), (short)0xF122);
        if(opt != null){
            try {
                EscherPropertyFactory f = new EscherPropertyFactory();
                List props = f.CreateProperties( opt.Serialize(), 8, opt.GetInstance() );
                EscherSimpleProperty p = (EscherSimpleProperty)props.Get(0);
                if(p.GetPropertyNumber() == 0x39F && p.GetPropertyValue() == 1){
                    group = new Table(spContainer, parent);
                } else {
                    group = new ShapeGroup(spContainer, parent);
                }
            } catch (Exception e){
                logger.log(POILogger.WARN, e.GetMessage());
                group = new ShapeGroup(spContainer, parent);
            }
        }  else {
            group = new ShapeGroup(spContainer, parent);
        }

        return group;
     }

    public static Shape CreateSimpeShape(EscherContainerRecord spContainer, Shape parent){
        Shape shape = null;
        EscherSpRecord spRecord = spContainer.GetChildById(EscherSpRecord.RECORD_ID);

        int type = spRecord.GetOptions() >> 4;
        switch (type){
            case ShapeTypes.TextBox:
                shape = new TextBox(spContainer, parent);
                break;
            case ShapeTypes.HostControl:
            case ShapeTypes.PictureFrame: {
                InteractiveInfo info = (InteractiveInfo)getClientDataRecord(spContainer, RecordTypes.InteractiveInfo.typeID);
                OEShapeAtom oes = (OEShapeAtom)getClientDataRecord(spContainer, RecordTypes.OEShapeAtom.typeID);
                if(info != null && info.GetInteractiveInfoAtom() != null){
                    switch(info.GetInteractiveInfoAtom().GetAction()){
                        case InteractiveInfoAtom.ACTION_OLE:
                            shape = new OLEShape(spContainer, parent);
                            break;
                        case InteractiveInfoAtom.ACTION_MEDIA:
                            shape = new MovieShape(spContainer, parent);
                            break;
                        default:
                            break;
                    }
                } else if (oes != null){
                    shape = new OLEShape(spContainer, parent);
                }

                if(shape == null) shape = new Picture(spContainer, parent);
                break;
            }
            case ShapeTypes.Line:
                shape = new Line(spContainer, parent);
                break;
            case ShapeTypes.NotPrimitive: {
                EscherOptRecord opt = (EscherOptRecord)Shape.GetEscherChild(spContainer, EscherOptRecord.RECORD_ID);
                EscherProperty prop = Shape.GetEscherProperty(opt, EscherProperties.GEOMETRY__VERTICES);
                if(prop != null)
                    shape = new Freeform(spContainer, parent);
                else {

                    logger.log(POILogger.WARN, "Creating AutoShape for a NotPrimitive shape");
                    shape = new AutoShape(spContainer, parent);
                }
                break;
            }
            default:
                shape = new AutoShape(spContainer, parent);
                break;
        }
        return shape;

    }

    protected static Record GetClientDataRecord(EscherContainerRecord spContainer, int recordType) {
        Record oep = null;
        for (Iterator<EscherRecord> it = spContainer.GetChildIterator(); it.HasNext();) {
            EscherRecord obj = it.next();
            if (obj.GetRecordId() == EscherClientDataRecord.RECORD_ID) {
                byte[] data = obj.Serialize();
                Record[] records = Record.FindChildRecords(data, 8, data.Length - 8);
                for (int j = 0; j < records.Length; j++) {
                    if (records[j].GetRecordType() == recordType) {
                        return records[j];
                    }
                }
            }
        }
        return oep;
    }

}





