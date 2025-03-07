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




using NPOI.ddf.EscherContainerRecord;
using NPOI.ddf.EscherDgRecord;
using NPOI.ddf.EscherDggRecord;
using NPOI.ddf.EscherSpRecord;
using NPOI.HSLF.record.ColorSchemeAtom;
using NPOI.HSLF.record.Comment2000;
using NPOI.HSLF.record.HeadersFootersContainer;
using NPOI.HSLF.record.Record;
using NPOI.HSLF.record.RecordContainer;
using NPOI.HSLF.record.RecordTypes;
using NPOI.HSLF.record.SlideAtom;
using NPOI.HSLF.record.TextHeaderAtom;
using NPOI.HSLF.record.SlideListWithText.SlideAtomsSet;

/**
 * This class represents a slide in a PowerPoint Document. It allows
 *  access to the text within, and the layout. For now, it only does
 *  the text side of things though
 *
 * @author Nick Burch
 * @author Yegor Kozlov
 */

public class Slide : Sheet
{
	private int _slideNo;
	private SlideAtomsSet _atomSet;
	private TextRun[] _Runs;
	private Notes _notes; // usermodel needs to Set this

	/**
	 * Constructs a Slide from the Slide record, and the SlideAtomsSet
	 *  Containing the text.
	 * Initialises TextRuns, to provide easier access to the text
	 *
	 * @param slide the Slide record we're based on
	 * @param notes the Notes sheet attached to us
	 * @param atomSet the SlideAtomsSet to Get the text from
	 */
	public Slide(NPOI.HSLF.record.Slide slide, Notes notes, SlideAtomsSet atomSet, int slideIdentifier, int slideNumber) {
        base(slide, slideIdentifier);

		_notes = notes;
		_atomSet = atomSet;
		_slideNo = slideNumber;

 		// Grab the TextRuns from the PPDrawing
		TextRun[] _otherRuns = FindTextRuns(getPPDrawing());

		// For the text coming in from the SlideAtomsSet:
		// Build up TextRuns from pairs of TextHeaderAtom and
		//  one of TextBytesAtom or TextCharsAtom
		Vector textRuns = new Vector();
		if(_atomSet != null) {
			FindTextRuns(_atomSet.GetSlideRecords(),textRuns);
		} else {
			// No text on the slide, must just be pictures
		}

		// Build an array, more useful than a vector
		_Runs = new TextRun[textRuns.Count+_otherRuns.Length];
		// Grab text from SlideListWithTexts entries
		int i=0;
		for(i=0; i<textRuns.Count; i++) {
			_Runs[i] = (TextRun)textRuns.Get(i);
            _Runs[i].SetSheet(this);
		}
		// Grab text from slide's PPDrawing
		for(int k=0; k<_otherRuns.Length; i++, k++) {
			_Runs[i] = _otherRuns[k];
            _Runs[i].SetSheet(this);
		}
	}

	/**
	* Create a new Slide instance
	* @param sheetNumber The internal number of the sheet, as used by PersistPtrHolder
	* @param slideNumber The user facing number of the sheet
	*/
	public Slide(int sheetNumber, int sheetRefId, int slideNumber){
		base(new NPOI.HSLF.record.Slide(), sheetNumber);
		_slideNo = slideNumber;
        GetSheetContainer().SetSheetId(sheetRefId);
	}

	/**
	 * Sets the Notes that are associated with this. Updates the
	 *  references in the records to point to the new ID
	 */
	public void SetNotes(Notes notes) {
		_notes = notes;

		// Update the Slide Atom's ID of where to point to
		SlideAtom sa = GetSlideRecord().GetSlideAtom();

		if(notes == null) {
			// Set to 0
			sa.SetNotesID(0);
		} else {
			// Set to the value from the notes' sheet id
			sa.SetNotesID(notes._getSheetNumber());
		}
	}

	/**
	* Changes the Slide's (external facing) page number.
	* @see NPOI.HSLF.usermodel.SlideShow#reorderSlide(int, int)
	*/
	public void SetSlideNumber(int newSlideNumber) {
		_slideNo = newSlideNumber;
	}

    /**
     * Called by SlideShow ater a new slide is Created.
     * <p>
     * For Slide we need to do the following:
     *  <li> Set id of the Drawing group.
     *  <li> Set shapeId for the Container descriptor and background
     * </p>
     */
    public void onCreate(){
        //Initialize Drawing group id
        EscherDggRecord dgg = GetSlideShow().GetDocumentRecord().GetPPDrawingGroup().GetEscherDggRecord();
        EscherContainerRecord dgContainer = (EscherContainerRecord)getSheetContainer().GetPPDrawing().GetEscherRecords()[0];
        EscherDgRecord dg = (EscherDgRecord) Shape.GetEscherChild(dgContainer, EscherDgRecord.RECORD_ID);
        int dgId = dgg.GetMaxDrawingGroupId() + 1;
        dg.SetOptions((short)(dgId << 4));
        dgg.SetDrawingsSaved(dgg.GetDrawingsSaved() + 1);
        dgg.SetMaxDrawingGroupId(dgId);

        foreach (EscherContainerRecord c in dgContainer.GetChildContainers()) {
            EscherSpRecord spr = null;
            switch(c.GetRecordId()){
                case EscherContainerRecord.SPGR_CONTAINER:
                    EscherContainerRecord dc = (EscherContainerRecord)c.GetChild(0);
                    spr = dc.GetChildById(EscherSpRecord.RECORD_ID);
                    break;
                case EscherContainerRecord.SP_CONTAINER:
                    spr = c.GetChildById(EscherSpRecord.RECORD_ID);
                    break;
            }
            if(spr != null) spr.SetShapeId(allocateShapeId());
        }

        //PPT doen't increment the number of saved shapes for group descriptor and background
        dg.SetNumShapes(1);
    }

	/**
	 * Create a <code>TextBox</code> object that represents the slide's title.
	 *
	 * @return <code>TextBox</code> object that represents the slide's title.
	 */
	public TextBox AddTitle() {
		Placeholder pl = new Placeholder();
		pl.SetShapeType(ShapeTypes.Rectangle);
		pl.GetTextRun().SetRunType(TextHeaderAtom.TITLE_TYPE);
		pl.SetText("Click to edit title");
		pl.SetAnchor(new java.awt.Rectangle(54, 48, 612, 90));
		AddShape(pl);
		return pl;
	}


	// Complex Accesser methods follow

	/**
	 * Return title of this slide or <code>null</code> if the slide does not have title.
	 * <p>
	 * The title is a run of text of type <code>TextHeaderAtom.CENTER_TITLE_TYPE</code> or
	 * <code>TextHeaderAtom.TITLE_TYPE</code>
	 * </p>
	 *
	 * @see TextHeaderAtom
	 *
	 * @return title of this slide
	 */
	public String GetTitle(){
		TextRun[] txt = GetTextRuns();
		for (int i = 0; i < txt.Length; i++) {
			int type = txt[i].GetRunType();
			if (type == TextHeaderAtom.CENTER_TITLE_TYPE ||
			type == TextHeaderAtom.TITLE_TYPE ){
				String title = txt[i].GetText();
				return title;
			}
		}
		return null;
	}

	// Simple Accesser methods follow

	/**
	 * Returns an array of all the TextRuns found
	 */
	public TextRun[] GetTextRuns() { return _Runs; }

	/**
	 * Returns the (public facing) page number of this slide
	 */
	public int GetSlideNumber() { return _slideNo; }

	/**
	 * Returns the underlying slide record
	 */
	public NPOI.HSLF.record.Slide GetSlideRecord() {
        return (NPOI.HSLF.record.Slide)getSheetContainer();
    }

	/**
	 * Returns the Notes Sheet for this slide, or null if there isn't one
	 */
	public Notes GetNotesSheet() { return _notes; }

	/**
	 * @return Set of records inside <code>SlideListWithtext</code> Container
	 *  which hold text data for this slide (typically for placeholders).
	 */
	protected SlideAtomsSet GetSlideAtomsSet() { return _atomSet;  }

    /**
     * Returns master sheet associated with this slide.
     * It can be either SlideMaster or TitleMaster objects.
     *
     * @return the master sheet associated with this slide.
     */
     public MasterSheet GetMasterSheet(){
        SlideMaster[] master = GetSlideShow().GetSlidesMasters();
        SlideAtom sa = GetSlideRecord().GetSlideAtom();
        int masterId = sa.GetMasterID();
        MasterSheet sheet = null;
        for (int i = 0; i < master.Length; i++) {
            if (masterId == master[i]._getSheetNumber()) {
                sheet = master[i];
                break;
            }
        }
        if (sheet == null){
            TitleMaster[] titleMaster = GetSlideShow().GetTitleMasters();
            if(titleMaster != null) for (int i = 0; i < titleMaster.Length; i++) {
                if (masterId == titleMaster[i]._getSheetNumber()) {
                    sheet = titleMaster[i];
                    break;
                }
            }
        }
        return sheet;
    }

    /**
     * Change Master of this slide.
     */
    public void SetMasterSheet(MasterSheet master){
        SlideAtom sa = GetSlideRecord().GetSlideAtom();
        int sheetNo = master._getSheetNumber();
        sa.SetMasterID(sheetNo);
    }

    /**
     * Sets whether this slide follows master background
     *
     * @param flag  <code>true</code> if the slide follows master,
     * <code>false</code> otherwise
     */
    public void SetFollowMasterBackground(bool flag){
        SlideAtom sa = GetSlideRecord().GetSlideAtom();
        sa.SetFollowMasterBackground(flag);
    }

    /**
     * Whether this slide follows master sheet background
     *
     * @return <code>true</code> if the slide follows master background,
     * <code>false</code> otherwise
     */
    public bool GetFollowMasterBackground(){
        SlideAtom sa = GetSlideRecord().GetSlideAtom();
        return sa.GetFollowMasterBackground();
    }

    /**
     * Sets whether this slide Draws master sheet objects
     *
     * @param flag  <code>true</code> if the slide Draws master sheet objects,
     * <code>false</code> otherwise
     */
    public void SetFollowMasterObjects(bool flag){
        SlideAtom sa = GetSlideRecord().GetSlideAtom();
        sa.SetFollowMasterObjects(flag);
    }

    /**
     * Whether this slide follows master color scheme
     *
     * @return <code>true</code> if the slide follows master color scheme,
     * <code>false</code> otherwise
     */
    public bool GetFollowMasterScheme(){
        SlideAtom sa = GetSlideRecord().GetSlideAtom();
        return sa.GetFollowMasterScheme();
    }

    /**
     * Sets whether this slide Draws master color scheme
     *
     * @param flag  <code>true</code> if the slide Draws master color scheme,
     * <code>false</code> otherwise
     */
    public void SetFollowMasterScheme(bool flag){
        SlideAtom sa = GetSlideRecord().GetSlideAtom();
        sa.SetFollowMasterScheme(flag);
    }

    /**
     * Whether this slide Draws master sheet objects
     *
     * @return <code>true</code> if the slide Draws master sheet objects,
     * <code>false</code> otherwise
     */
    public bool GetFollowMasterObjects(){
        SlideAtom sa = GetSlideRecord().GetSlideAtom();
        return sa.GetFollowMasterObjects();
    }

    /**
     * Background for this slide.
     */
     public Background GetBackground() {
        if(getFollowMasterBackground()) {
            return GetMasterSheet().GetBackground();
        }
        return super.GetBackground();
    }

    /**
     * Color scheme for this slide.
     */
    public ColorSchemeAtom GetColorScheme() {
        if(getFollowMasterScheme()){
            return GetMasterSheet().GetColorScheme();
        }
        return super.GetColorScheme();
    }

    /**
     * Get the comment(s) for this slide.
     * Note - for now, only works on PPT 2000 and
     *  PPT 2003 files. Doesn't work for PPT 97
     *  ones, as they do their comments oddly.
     */
    public Comment[] GetComments() {
    	// If there are any, they're in
    	//  ProgTags -> ProgBinaryTag -> BinaryTagData
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
    			RecordContainer binaryTags = (RecordContainer)
    				progBinaryTag.FindFirstOfType(
    						RecordTypes.BinaryTagData.typeID
    			);
    			if(binaryTags != null) {
    				// This is where they'll be
    				int count = 0;
    				for(int i=0; i<binaryTags.GetChildRecords().Length; i++) {
    					if(binaryTags.GetChildRecords()[i] is Comment2000) {
    						count++;
    					}
    				}

    				// Now build
    				Comment[] comments = new Comment[count];
    				count = 0;
    				for(int i=0; i<binaryTags.GetChildRecords().Length; i++) {
    					if(binaryTags.GetChildRecords()[i] is Comment2000) {
    						comments[i] = new Comment(
    								(Comment2000)binaryTags.GetChildRecords()[i]
    						);
    						count++;
    					}
    				}

    				return comments;
    			}
    		}
    	}

    	// None found
    	return Array.Empty<Comment>();
    }

    public void Draw(Graphics2D graphics){
        MasterSheet master = GetMasterSheet();
        if(getFollowMasterBackground()) master.GetBackground().Draw(graphics);
        if(getFollowMasterObjects()){
            Shape[] sh = master.GetShapes();
            for (int i = 0; i < sh.Length; i++) {
                if(MasterSheet.IsPlaceholder(sh[i])) continue;

                sh[i].Draw(graphics);
            }
        }
        Shape[] sh = GetShapes();
        for (int i = 0; i < sh.Length; i++) {
            sh[i].Draw(graphics);
        }
    }

    /**
     * Header / Footer Settings for this slide.
     *
     * @return Header / Footer Settings for this slide
     */
     public HeadersFooters GetHeadersFooters(){
        HeadersFootersContainer hdd = null;
        Record[] ch = GetSheetContainer().GetChildRecords();
        bool ppt2007 = false;
        for (int i = 0; i < ch.Length; i++) {
            if(ch[i] is HeadersFootersContainer){
                hdd = (HeadersFootersContainer)ch[i];
            } else if (ch[i].GetRecordType() == RecordTypes.RoundTripContentMasterId.typeID){
                ppt2007 = true;
            }
        }
        bool newRecord = false;
        if(hdd == null && !ppt2007) {
            return GetSlideShow().GetSlideHeadersFooters();
        }
        if(hdd == null) {
            hdd = new HeadersFootersContainer(HeadersFootersContainer.SlideHeadersFootersContainer);
            newRecord = true;
        }
        return new HeadersFooters(hdd, this, newRecord, ppt2007);
    }

    protected void onAddTextShape(TextShape shape) {
        TextRun run = shape.GetTextRun();

        if(_Runs == null) _Runs = new TextRun[]{Run};
        else {
            TextRun[] tmp = new TextRun[_Runs.Length + 1];
            Array.Copy(_Runs, 0, tmp, 0, _Runs.Length);
            tmp[tmp.Length-1] = Run;
            _Runs = tmp;
        }
    }
}





