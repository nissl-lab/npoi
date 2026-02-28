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



using NPOI.ddf.EscherClientDataRecord;
using NPOI.ddf.EscherContainerRecord;
using NPOI.ddf.EscherProperties;
using NPOI.HSLF.exceptions.HSLFException;
using NPOI.HSLF.record.*;
using NPOI.HSLF.usermodel.SlideShow;

/**
 * Represents a movie in a PowerPoint document.
 *
 * @author Yegor Kozlov
 */
public class MovieShape : Picture {
    public static int DEFAULT_MOVIE_THUMBNAIL = -1;

    public static int MOVIE_MPEG = 1;
    public static int MOVIE_AVI  = 2;

    /**
     * Create a new <code>Picture</code>
     *
    * @param pictureIdx the index of the picture
     */
    public MovieShape(int movieIdx, int pictureIdx){
        base(pictureIdx, null);
        SetMovieIndex(movieIdx);
        SetAutoPlay(true);
    }

    /**
     * Create a new <code>Picture</code>
     *
     * @param idx the index of the picture
     * @param parent the parent shape
     */
    public MovieShape(int movieIdx, int idx, Shape parent) {
        base(idx, parent);
        SetMovieIndex(movieIdx);
    }

    /**
      * Create a <code>Picture</code> object
      *
      * @param escherRecord the <code>EscherSpContainer</code> record which holds information about
      *        this picture in the <code>Slide</code>
      * @param parent the parent shape of this picture
      */
     protected MovieShape(EscherContainerRecord escherRecord, Shape parent){
        base(escherRecord, parent);
    }

    /**
     * Create a new Placeholder and Initialize internal structures
     *
     * @return the Created <code>EscherContainerRecord</code> which holds shape data
     */
    protected EscherContainerRecord CreateSpContainer(int idx, bool IsChild) {
        _escherContainer = super.CreateSpContainer(idx, isChild);

        SetEscherProperty(EscherProperties.PROTECTION__LOCKAGAINSTGROUPING, 0x1000100);
        SetEscherProperty(EscherProperties.FILL__NOFILLHITTEST, 0x10001);

        EscherClientDataRecord cldata = new EscherClientDataRecord();
        cldata.SetOptions((short)0xF);
        _escherContainer.AddChildRecord(cldata);

        OEShapeAtom oe = new OEShapeAtom();
        InteractiveInfo info = new InteractiveInfo();
        InteractiveInfoAtom infoAtom = info.GetInteractiveInfoAtom();
        infoAtom.SetAction(InteractiveInfoAtom.ACTION_MEDIA);
        infoAtom.SetHyperlinkType(InteractiveInfoAtom.LINK_NULL);

        AnimationInfo an = new AnimationInfo();
        AnimationInfoAtom anAtom = an.GetAnimationInfoAtom();
        anAtom.SetFlag(AnimationInfoAtom.Automatic, true);

        //convert hslf into ddf
        MemoryStream out = new MemoryStream();
        try {
            oe.WriteOut(out);
            an.WriteOut(out);
            info.WriteOut(out);
        } catch(Exception e){
            throw new HSLFException(e);
        }
        cldata.SetRemainingData(out.ToArray());

        return _escherContainer;
    }

    /**
     * Assign a movie to this shape
     *
     * @see NPOI.HSLF.usermodel.SlideShow#AddMovie(String, int)
     * @param idx  the index of the movie
     */
    public void SetMovieIndex(int idx){
        OEShapeAtom oe = (OEShapeAtom)getClientDataRecord(RecordTypes.OEShapeAtom.typeID);
        oe.SetOptions(idx);

        AnimationInfo an = (AnimationInfo)getClientDataRecord(RecordTypes.AnimationInfo.typeID);
        if(an != null) {
            AnimationInfoAtom ai = an.GetAnimationInfoAtom();
            ai.SetDimColor(0x07000000);
            ai.SetFlag(AnimationInfoAtom.Automatic, true);
            ai.SetFlag(AnimationInfoAtom.Play, true);
            ai.SetFlag(AnimationInfoAtom.Synchronous, true);
            ai.SetOrderID(idx + 1);
        }
    }

    public void SetAutoPlay(bool flag){
        AnimationInfo an = (AnimationInfo)getClientDataRecord(RecordTypes.AnimationInfo.typeID);
        if(an != null){
            an.GetAnimationInfoAtom().SetFlag(AnimationInfoAtom.Automatic, flag);
            updateClientData();
        }
    }

    public bool  isAutoPlay(){
        AnimationInfo an = (AnimationInfo)getClientDataRecord(RecordTypes.AnimationInfo.typeID);
        if(an != null){
            return an.GetAnimationInfoAtom().GetFlag(AnimationInfoAtom.Automatic);
        }
        return false;
    }

    /**
     * @return UNC or local path to a video file
     */
    public String GetPath(){
        OEShapeAtom oe = (OEShapeAtom)getClientDataRecord(RecordTypes.OEShapeAtom.typeID);
        int idx = oe.GetOptions();

        SlideShow ppt = Sheet.GetSlideShow();
        ExObjList lst = (ExObjList)ppt.GetDocumentRecord().FindFirstOfType(RecordTypes.ExObjList.typeID);
        if(lst == null) return null;

        Record[]  r = lst.GetChildRecords();
        for (int i = 0; i < r.Length; i++) {
            if(r[i] is ExMCIMovie){
                ExMCIMovie mci = (ExMCIMovie)r[i];
                ExVideoContainer exVideo = mci.GetExVideo();
                int objectId = exVideo.GetExMediaAtom().GetObjectId();
                if(objectId == idx){
                    return exVideo.GetPathAtom().GetText();
                }
            }

        }
        return null;
    }
}





