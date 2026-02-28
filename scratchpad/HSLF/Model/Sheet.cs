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
using NPOI.HSLF.usermodel.SlideShow;







/**
 * This class defines the common format of "Sheets" in a powerpoint
 * document. Such sheets could be Slides, Notes, Master etc
 *
 * @author Nick Burch
 * @author Yegor Kozlov
 */

public abstract class Sheet {
    /**
     * The <code>SlideShow</code> we belong to
     */
    private SlideShow _slideShow;

    /**
     * Sheet background
     */
    private Background _background;

    /**
     * Record Container that holds sheet data.
     * For slides it is NPOI.HSLF.record.Slide,
     * for notes it is NPOI.HSLF.record.Notes,
     * for slide masters it is NPOI.HSLF.record.SlideMaster, etc.
     */
    private SheetContainer _Container;

    private int _sheetNo;

    public Sheet(SheetContainer Container, int sheetNo) {
        _Container = Container;
        _sheetNo = sheetNo;
    }

    /**
     * Returns an array of all the TextRuns in the sheet.
     */
    public abstract TextRun[] GetTextRuns();

    /**
     * Returns the (internal, RefID based) sheet number, as used
     * to in PersistPtr stuff.
     */
    public int _getSheetRefId() {
        return _Container.GetSheetId();
    }

    /**
     * Returns the (internal, SlideIdentifier based) sheet number, as used
     * to reference this sheet from other records.
     */
    public int _getSheetNumber() {
        return _sheetNo;
    }

    /**
     * Fetch the PPDrawing from the underlying record
     */
    protected PPDrawing GetPPDrawing() {
        return _Container.GetPPDrawing();
    }

    /**
     * Fetch the SlideShow we're attached to
     */
    public SlideShow GetSlideShow() {
        return _slideShow;
    }

    /**
     * Return record Container for this sheet
     */
    public SheetContainer GetSheetContainer() {
        return _Container;
    }

    /**
     * Set the SlideShow we're attached to.
     * Also passes it on to our child RichTextRuns
     */
    public void SetSlideShow(SlideShow ss) {
        _slideShow = ss;
        TextRun[] trs = GetTextRuns();
        if (trs != null) {
            for (int i = 0; i < trs.Length; i++) {
                trs[i].supplySlideShow(_slideShow);
            }
        }
    }


    /**
     * For a given PPDrawing, grab all the TextRuns
     */
    public static TextRun[] FindTextRuns(PPDrawing ppdrawing) {
        Vector RunsV = new Vector();
        EscherTextboxWrapper[] wrappers = ppdrawing.GetTextboxWrappers();
        for (int i = 0; i < wrappers.Length; i++) {
            int s1 = RunsV.Count;

            // propagate parents to parent-aware records
            RecordContainer.handleParentAwareRecords(wrappers[i]);
            FindTextRuns(wrappers[i].GetChildRecords(), RunsV);
            int s2 = RunsV.Count;
            if (s2 != s1){
                TextRun t = (TextRun) RunsV.Get(RunsV.Count-1);
                t.SetShapeId(wrappers[i].GetShapeId());
            }
        }
        TextRun[] Runs = new TextRun[RunsV.Count];
        for (int i = 0; i < Runs.Length; i++) {
            Runs[i] = (TextRun) RunsV.Get(i);
        }
        return Runs;
    }

    /**
     * Scans through the supplied record array, looking for
     * a TextHeaderAtom followed by one of a TextBytesAtom or
     * a TextCharsAtom. Builds up TextRuns from these
     *
     * @param records the records to build from
     * @param found   vector to add any found to
     */
    protected static void FindTextRuns(Record[] records, Vector found) {
        // Look for a TextHeaderAtom
        for (int i = 0, slwtIndex=0; i < (records.Length - 1); i++) {
            if (records[i] is TextHeaderAtom) {
                TextRun trun = null;
                TextHeaderAtom tha = (TextHeaderAtom) records[i];
                StyleTextPropAtom stpa = null;

                // Look for a subsequent StyleTextPropAtom
                if (i < (records.Length - 2)) {
                    if (records[i + 2] is StyleTextPropAtom) {
                        stpa = (StyleTextPropAtom) records[i + 2];
                    }
                }

                // See what follows the TextHeaderAtom
                if (records[i + 1] is TextCharsAtom) {
                    TextCharsAtom tca = (TextCharsAtom) records[i + 1];
                    trun = new TextRun(tha, tca, stpa);
                } else if (records[i + 1] is TextBytesAtom) {
                    TextBytesAtom tba = (TextBytesAtom) records[i + 1];
                    trun = new TextRun(tha, tba, stpa);
                } else if (records[i + 1].GetRecordType() == 4001l) {
                    // StyleTextPropAtom - Safe to ignore
                } else if (records[i + 1].GetRecordType() == 4010l) {
                    // TextSpecInfoAtom - Safe to ignore
                } else {
                    System.err.println("Found a TextHeaderAtom not followed by a TextBytesAtom or TextCharsAtom: Followed by " + records[i + 1].GetRecordType());
                }

                if (trun != null) {
                    ArrayList lst = new ArrayList();
                    for (int j = i; j < records.Length; j++) {
                        if(j > i && records[j] is TextHeaderAtom) break;
                        lst.Add(records[j]);
                    }
                    Record[] recs = new Record[lst.Count];
                    lst.ToArray(recs);
                    tRun._records = recs;
                    tRun.SetIndex(slwtIndex);

                    found.Add(tRun);
                    i++;
                } else {
                    // Not a valid one, so skip on to next and look again
                }
                slwtIndex++;
            }
        }
    }

    /**
     * Returns all shapes Contained in this Sheet
     *
     * @return all shapes Contained in this Sheet (Slide or Notes)
     */
    public Shape[] GetShapes() {
        PPDrawing ppdrawing = GetPPDrawing();

        EscherContainerRecord dg = (EscherContainerRecord) ppdrawing.GetEscherRecords()[0];
        EscherContainerRecord spgr = null;

        for (Iterator<EscherRecord> it = dg.GetChildIterator(); it.HasNext();) {
            EscherRecord rec = it.next();
            if (rec.GetRecordId() == EscherContainerRecord.SPGR_CONTAINER) {
                spgr = (EscherContainerRecord) rec;
                break;
            }
        }
        if (spgr == null) {
            throw new InvalidOperationException("spgr not found");
        }

        List<Shape> shapes = new List<Shape>();
        Iterator<EscherRecord> it = spgr.GetChildIterator();
        if (it.HasNext()) {
            // skip first item
            it.next();
        }
        for (; it.HasNext();) {
            EscherContainerRecord sp = (EscherContainerRecord) it.next();
            Shape sh = ShapeFactory.CreateShape(sp, null);
            sh.SetSheet(this);
            shapes.Add(sh);
        }

        return shapes.ToArray(new Shape[shapes.Count]);
    }

    /**
     * Add a new Shape to this Slide
     *
     * @param shape - the Shape to add
     */
    public void AddShape(Shape shape) {
        PPDrawing ppdrawing = GetPPDrawing();

        EscherContainerRecord dgContainer = (EscherContainerRecord) ppdrawing.GetEscherRecords()[0];
        EscherContainerRecord spgr = (EscherContainerRecord) Shape.GetEscherChild(dgContainer, EscherContainerRecord.SPGR_CONTAINER);
        spgr.AddChildRecord(shape.GetSpContainer());

        shape.SetSheet(this);
        shape.SetShapeId(allocateShapeId());
        shape.afterInsert(this);
    }

    /**
     * Allocates new shape id for the new Drawing group id.
     *
     * @return a new shape id.
     */
    public int allocateShapeId()
    {
        EscherDggRecord dgg = _slideShow.GetDocumentRecord().GetPPDrawingGroup().GetEscherDggRecord();
        EscherDgRecord dg = _Container.GetPPDrawing().GetEscherDgRecord();

        dgg.SetNumShapesSaved( dgg.GetNumShapesSaved() + 1 );

        // Add to existing cluster if space available
        for (int i = 0; i < dgg.GetFileIdClusters().Length; i++)
        {
            EscherDggRecord.FileIdCluster c = dgg.GetFileIdClusters()[i];
            if (c.GetDrawingGroupId() == dg.GetDrawingGroupId() && c.GetNumShapeIdsUsed() != 1024)
            {
                int result = c.GetNumShapeIdsUsed() + (1024 * (i+1));
                c.incrementShapeId();
                dg.SetNumShapes( dg.GetNumShapes() + 1 );
                dg.SetLastMSOSPID( result );
                if (result >= dgg.GetShapeIdMax())
                    dgg.SetShapeIdMax( result + 1 );
                return result;
            }
        }

        // Create new cluster
        dgg.AddCluster( dg.GetDrawingGroupId(), 0, false );
        dgg.GetFileIdClusters()[dgg.GetFileIdClusters().Length-1].incrementShapeId();
        dg.SetNumShapes( dg.GetNumShapes() + 1 );
        int result = (1024 * dgg.GetFileIdClusters().Length);
        dg.SetLastMSOSPID( result );
        if (result >= dgg.GetShapeIdMax())
            dgg.SetShapeIdMax( result + 1 );
        return result;
    }

    /**
     * Removes the specified shape from this sheet.
     *
     * @param shape shape to be Removed from this sheet, if present.
     * @return <tt>true</tt> if the shape was deleted.
     */
    public bool RemoveShape(Shape shape) {
        PPDrawing ppdrawing = GetPPDrawing();

        EscherContainerRecord dg = (EscherContainerRecord) ppdrawing.GetEscherRecords()[0];
        EscherContainerRecord spgr = null;

        for (Iterator<EscherRecord> it = dg.GetChildIterator(); it.HasNext();) {
            EscherRecord rec = it.next();
            if (rec.GetRecordId() == EscherContainerRecord.SPGR_CONTAINER) {
                spgr = (EscherContainerRecord) rec;
                break;
            }
        }
        if(spgr == null) {
            return false;
        }

        List<EscherRecord> lst = spgr.GetChildRecords();
        bool result = lst.Remove(shape.GetSpContainer());
        spgr.SetChildRecords(lst);
        return result;
    }

    /**
     * Called by SlideShow ater a new sheet is Created
     */
    public void onCreate(){

    }

    /**
     * Return the master sheet .
     */
    public abstract MasterSheet GetMasterSheet();

    /**
     * Color scheme for this sheet.
     */
    public ColorSchemeAtom GetColorScheme() {
        return _Container.GetColorScheme();
    }

    /**
     * Returns the background shape for this sheet.
     *
     * @return the background shape for this sheet.
     */
    public Background GetBackground() {
        if (_background == null) {
            PPDrawing ppdrawing = GetPPDrawing();

            EscherContainerRecord dg = (EscherContainerRecord) ppdrawing.GetEscherRecords()[0];
            EscherContainerRecord spContainer = null;

            for (Iterator<EscherRecord> it = dg.GetChildIterator(); it.HasNext();) {
                EscherRecord rec = it.next();
                if (rec.GetRecordId() == EscherContainerRecord.SP_CONTAINER) {
                    spContainer = (EscherContainerRecord) rec;
                    break;
                }
            }
            _background = new Background(spContainer, null);
            _background.SetSheet(this);
        }
        return _background;
    }

    public void Draw(Graphics2D graphics){

    }

    /**
     * Subclasses should call this method and update the array of text Runs
     * when a text shape is Added
     *
     * @param shape
     */
    protected void onAddTextShape(TextShape shape) {

    }

    /**
     * Return placeholder by text type
     *
     * @param type  type of text, See {@link NPOI.HSLF.record.TextHeaderAtom}
     * @return  <code>TextShape</code> or <code>null</code>
     */
    public TextShape GetPlaceholderByTextType(int type){
        Shape[] shape = GetShapes();
        for (int i = 0; i < shape.Length; i++) {
            if(shape[i] is TextShape){
                TextShape tx = (TextShape)shape[i];
                TextRun run = tx.GetTextRun();
                if(run != null && Run.GetRunType() == type){
                    return tx;
                }
            }
        }
        return null;
    }

    /**
     * Search text placeholer by its type
     *
     * @param type  type of placeholder to search. See {@link NPOI.HSLF.record.OEPlaceholderAtom}
     * @return  <code>TextShape</code> or <code>null</code>
     */
    public TextShape GetPlaceholder(int type){
        Shape[] shape = GetShapes();
        for (int i = 0; i < shape.Length; i++) {
            if(shape[i] is TextShape){
                TextShape tx = (TextShape)shape[i];
                int placeholderId = 0;
                OEPlaceholderAtom oep = tx.GetPlaceholderAtom();
                if(oep != null) {
                    placeholderId = oep.GetPlaceholderId();
                } else {
                    //special case for files saved in Office 2007
                    RoundTripHFPlaceholder12 hldr = (RoundTripHFPlaceholder12)tx.GetClientDataRecord(RecordTypes.RoundTripHFPlaceholder12.typeID);
                    if(hldr != null) placeholderId = hldr.GetPlaceholderId();
                }
                if(placeholderId == type){
                    return tx;
                }
            }
        }
        return null;
    }

    /**
     * Return programmable tag associated with this sheet, e.g. <code>___PPT12</code>.
     *
     * @return programmable tag associated with this sheet.
     */
    public String GetProgrammableTag(){
        String tag = null;
        RecordContainer progTags = (RecordContainer)
                GetSheetContainer().FindFirstOfType(
                            RecordTypes.ProgTags.typeID
        );
        if(progTags != null) {
            RecordContainer progBinaryTag = (RecordContainer)
                progTags.FindFirstOfType(
                        RecordTypes.ProgBinaryTag.typeID
            );
            if(progBinaryTag != null) {
                CString binaryTag = (CString)
                    progBinaryTag.FindFirstOfType(
                            RecordTypes.CString.typeID
                );
                if(binaryTag != null) tag = binaryTag.GetText();
            }
        }

        return tag;

    }

}





