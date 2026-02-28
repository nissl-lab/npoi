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
using NPOI.HSLF.record.*;
using NPOI.HSLF.exceptions.HSLFException;
using NPOI.util.LittleEndian;





/**
 * Represents an ActiveX control in a PowerPoint document.
 *
 * TODO: finish
 * @author Yegor Kozlov
 */
public class ActiveXShape : Picture {
    public static int DEFAULT_ACTIVEX_THUMBNAIL = -1;

    /**
     * Create a new <code>Picture</code>
     *
    * @param pictureIdx the index of the picture
     */
    public ActiveXShape(int movieIdx, int pictureIdx){
        base(pictureIdx, null);
        SetActiveXIndex(movieIdx);
    }

    /**
      * Create a <code>Picture</code> object
      *
      * @param escherRecord the <code>EscherSpContainer</code> record which holds information about
      *        this picture in the <code>Slide</code>
      * @param parent the parent shape of this picture
      */
     protected ActiveXShape(EscherContainerRecord escherRecord, Shape parent){
        base(escherRecord, parent);
    }

    /**
     * Create a new Placeholder and Initialize internal structures
     *
     * @return the Created <code>EscherContainerRecord</code> which holds shape data
     */
    protected EscherContainerRecord CreateSpContainer(int idx, bool IsChild) {
        _escherContainer = super.CreateSpContainer(idx, isChild);

        EscherSpRecord spRecord = _escherContainer.GetChildById(EscherSpRecord.RECORD_ID);
        spRecord.SetFlags(EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE | EscherSpRecord.FLAG_OLESHAPE);

        SetShapeType(ShapeTypes.HostControl);
        SetEscherProperty(EscherProperties.BLIP__PICTUREID, idx);
        SetEscherProperty(EscherProperties.LINESTYLE__COLOR, 0x8000001);
        SetEscherProperty(EscherProperties.LINESTYLE__NOLINEDRAWDASH, 0x80008);
        SetEscherProperty(EscherProperties.SHADOWSTYLE__COLOR, 0x8000002);
        SetEscherProperty(EscherProperties.PROTECTION__LOCKAGAINSTGROUPING, -1);

        EscherClientDataRecord cldata = new EscherClientDataRecord();
        cldata.SetOptions((short)0xF);
        _escherContainer.AddChildRecord(cldata); // TODO unit Test to prove GetChildRecords().add is wrong

        OEShapeAtom oe = new OEShapeAtom();

        //convert hslf into ddf
        MemoryStream out = new MemoryStream();
        try {
            oe.WriteOut(out);
        } catch(Exception e){
            throw new HSLFException(e);
        }
        cldata.SetRemainingData(out.ToArray());

        return _escherContainer;
    }

    /**
     * Assign a control to this shape
     *
     * @see NPOI.HSLF.usermodel.SlideShow#AddMovie(String, int)
     * @param idx  the index of the movie
     */
    public void SetActiveXIndex(int idx){
        EscherContainerRecord spContainer = GetSpContainer();
        for (Iterator<EscherRecord> it = spContainer.GetChildIterator(); it.HasNext();) {
            EscherRecord obj = it.next();
            if (obj.GetRecordId() == EscherClientDataRecord.RECORD_ID) {
                EscherClientDataRecord clientRecord = (EscherClientDataRecord)obj;
                byte[] recdata = clientRecord.GetRemainingData();
                LittleEndian.PutInt(recdata, 8, idx);
            }
        }
    }

    public int GetControlIndex(){
        int idx = -1;
        OEShapeAtom oe = (OEShapeAtom)getClientDataRecord(RecordTypes.OEShapeAtom.typeID);
        if(oe != null) idx = oe.GetOptions();
        return idx;
    }

    /**
     * Set a property of this ActiveX control
     * @param key
     * @param value
     */
    public void SetProperty(String key, String value){

    }

    /**
     * Document-level Container that specifies information about an ActiveX control
     *
     * @return Container that specifies information about an ActiveX control
     */
    public ExControl GetExControl(){
        int idx = GetControlIndex();
        ExControl ctrl = null;
        Document doc = Sheet.GetSlideShow().GetDocumentRecord();
        ExObjList lst = (ExObjList)doc.FindFirstOfType(RecordTypes.ExObjList.typeID);
        if(lst != null){
            Record[] ch = lst.GetChildRecords();
            for (int i = 0; i < ch.Length; i++) {
                if(ch[i] is ExControl){
                    ExControl c = (ExControl)ch[i];
                    if(c.GetExOleObjAtom().GetObjID() == idx){
                        ctrl = c;
                        break;
                    }
                }
            }
        }
        return ctrl;
    }

    protected void afterInsert(Sheet sheet){
        ExControl ctrl = GetExControl();
        ctrl.GetExControlAtom().SetSlideId(sheet._getSheetNumber());

        try {
            String name = ctrl.GetProgId() + "-" + GetControlIndex();
            byte[] data = (name + '\u0000').GetBytes("UTF-16LE");
            EscherComplexProperty prop = new EscherComplexProperty(EscherProperties.GROUPSHAPE__SHAPENAME, false, data);
            EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
            opt.AddEscherProperty(prop);
        } catch (UnsupportedEncodingException e){
            throw new HSLFException(e);
        }
    }
}





