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




using NPOI.HSLF.Model.textproperties.TextPropCollection;
using NPOI.HSLF.record.*;
using NPOI.HSLF.usermodel.RichTextRun;
using NPOI.HSLF.usermodel.SlideShow;
using NPOI.util.StringUtil;

/**
 * This class represents a run of text in a powerpoint document. That
 *  run could be text on a sheet, or text in a note.
 *  It is only a very basic class for now
 *
 * @author Nick Burch
 */

public class TextRun
{
	// Note: These fields are protected to help with unit Testing
	//   Other classes shouldn't really go playing with them!
	protected TextHeaderAtom _headerAtom;
	protected TextBytesAtom  _byteAtom;
	protected TextCharsAtom  _charAtom;
	protected StyleTextPropAtom _styleAtom;
    protected TextRulerAtom _ruler;
    protected bool _isUnicode;
	protected RichTextRun[] _rtRuns;
	private SlideShow slideShow;
    private Sheet _sheet;
    private int shapeId;
    private int slwtIndex; //position in the owning SlideListWithText
    /**
     * all text run records that follow TextHeaderAtom.
     * (there can be misc InteractiveInfo, TxInteractiveInfo and other records)
     */
    protected Record[] _records;

	/**
	* Constructs a Text Run from a Unicode text block
	*
	* @param tha the TextHeaderAtom that defines what's what
	* @param tca the TextCharsAtom Containing the text
	* @param sta the StyleTextPropAtom which defines the character stylings
	*/
	public TextRun(TextHeaderAtom tha, TextCharsAtom tca, StyleTextPropAtom sta) {
		this(tha,null,tca,sta);
	}

	/**
	* Constructs a Text Run from a Ascii text block
	*
	* @param tha the TextHeaderAtom that defines what's what
	* @param tba the TextBytesAtom Containing the text
	* @param sta the StyleTextPropAtom which defines the character stylings
	*/
	public TextRun(TextHeaderAtom tha, TextBytesAtom tba, StyleTextPropAtom sta) {
		this(tha,tba,null,sta);
	}

	/**
	 * Internal constructor and Initializer
	 */
	private TextRun(TextHeaderAtom tha, TextBytesAtom tba, TextCharsAtom tca, StyleTextPropAtom sta) {
		_headerAtom = tha;
		_styleAtom = sta;
		if(tba != null) {
			_byteAtom = tba;
			_isUnicode = false;
		} else {
			_charAtom = tca;
			_isUnicode = true;
		}
		String RunRawText = GetText();

		// Figure out the rich text Runs
		LinkedList pStyles = new LinkedList();
		LinkedList cStyles = new LinkedList();
		if(_styleAtom != null) {
			// Get the style atom to grok itself
			_styleAtom.SetParentTextSize(RunRawText.Length);
			pStyles = _styleAtom.GetParagraphStyles();
			cStyles = _styleAtom.GetCharacterStyles();
		}
        buildRichTextRuns(pStyles, cStyles, RunRawText);
	}

	public void buildRichTextRuns(LinkedList pStyles, LinkedList cStyles, String RunRawText){

        // Handle case of no current style, with a default
        if(pStyles.Count == 0 || cStyles.Count == 0) {
            _rtRuns = new RichTextRun[1];
            _rtRuns[0] = new RichTextRun(this, 0, RunRawText.Length);
        } else {
            // Build up Rich Text Runs, one for each
            //  character/paragraph style pair
            Vector rtrs = new Vector();

            int pos = 0;

            int curP = 0;
            int curC = 0;
            int pLenRemain = -1;
            int cLenRemain = -1;

            // Build one for each run with the same style
            while(pos <= RunRawText.Length && curP < pStyles.Count && curC < cStyles.Count) {
                // Get the Props to use
                TextPropCollection pProps = (TextPropCollection)pStyles.Get(curP);
                TextPropCollection cProps = (TextPropCollection)cStyles.Get(curC);

                int pLen = pProps.GetCharactersCovered();
                int cLen = cProps.GetCharactersCovered();

                // Handle new pass
                bool freshSet = false;
                if(pLenRemain == -1 && cLenRemain == -1) { freshSet = true; }
                if(pLenRemain == -1) { pLenRemain = pLen; }
                if(cLenRemain == -1) { cLenRemain = cLen; }

                // So we know how to build the eventual run
                int RunLen = -1;
                bool pShared = false;
                bool cShared = false;

                // Same size, new styles - neither shared
                if(pLen == cLen && freshSet) {
                    RunLen = cLen;
                    pShared = false;
                    cShared = false;
                    curP++;
                    curC++;
                    pLenRemain = -1;
                    cLenRemain = -1;
                } else {
                    // Some sharing

                    // See if we are already in a shared block
                    if(pLenRemain < pLen) {
                        // Existing shared p block
                        pShared = true;

                        // Do we end with the c block, or either side of it?
                        if(pLenRemain == cLenRemain) {
                            // We end at the same time
                            cShared = false;
                            RunLen = pLenRemain;
                            curP++;
                            curC++;
                            pLenRemain = -1;
                            cLenRemain = -1;
                        } else if(pLenRemain < cLenRemain) {
                            // We end before the c block
                            cShared = true;
                            RunLen = pLenRemain;
                            curP++;
                            cLenRemain -= pLenRemain;
                            pLenRemain = -1;
                        } else {
                            // We end after the c block
                            cShared = false;
                            RunLen = cLenRemain;
                            curC++;
                            pLenRemain -= cLenRemain;
                            cLenRemain = -1;
                        }
                    } else if(cLenRemain < cLen) {
                        // Existing shared c block
                        cShared = true;

                        // Do we end with the p block, or either side of it?
                        if(pLenRemain == cLenRemain) {
                            // We end at the same time
                            pShared = false;
                            RunLen = cLenRemain;
                            curP++;
                            curC++;
                            pLenRemain = -1;
                            cLenRemain = -1;
                        } else if(cLenRemain < pLenRemain) {
                            // We end before the p block
                            pShared = true;
                            RunLen = cLenRemain;
                            curC++;
                            pLenRemain -= cLenRemain;
                            cLenRemain = -1;
                        } else {
                            // We end after the p block
                            pShared = false;
                            RunLen = pLenRemain;
                            curP++;
                            cLenRemain -= pLenRemain;
                            pLenRemain = -1;
                        }
                    } else {
                        // Start of a shared block
                        if(pLenRemain < cLenRemain) {
                            // Shared c block
                            pShared = false;
                            cShared = true;
                            RunLen = pLenRemain;
                            curP++;
                            cLenRemain -= pLenRemain;
                            pLenRemain = -1;
                        } else {
                            // Shared p block
                            pShared = true;
                            cShared = false;
                            RunLen = cLenRemain;
                            curC++;
                            pLenRemain -= cLenRemain;
                            cLenRemain = -1;
                        }
                    }
                }

                // Wind on
                int prevPos = pos;
                pos += RunLen;
                // Adjust for end-of-run extra 1 length
                if(pos > RunRawText.Length) {
                    RunLen--;
                }

                // Save
                RichTextRun rtr = new RichTextRun(this, prevPos, RunLen, pProps, cProps, pShared, cShared);
                rtrs.Add(rtr);
            }

            // Build the array
            _rtRuns = new RichTextRun[rtrs.Count];
            rtrs.copyInto(_rtRuns);
        }

    }

    // Update methods follow

	/**
	 * Adds the supplied text onto the end of the TextRun,
	 *  creating a new RichTextRun (returned) for it to
	 *  sit in.
	 * In many cases, before calling this, you'll want to add
	 *  a newline onto the end of your last RichTextRun
	 */
	public RichTextRun AppendText(String s) {
		// We will need a StyleTextProp atom
		ensureStyleAtomPresent();

		// First up, append the text to the
		//  underlying text atom
		int oldSize = GetRawText().Length;
		storeText(
			 GetRawText() + s
		);

		// If either of the previous styles overran
		//  the text by one, we need to shuffle that
		//  extra character onto the new ones
		int pOverRun = _styleAtom.GetParagraphTextLengthCovered() - oldSize;
		int cOverRun = _styleAtom.GetCharacterTextLengthCovered() - oldSize;
		if(pOverRun > 0) {
			TextPropCollection tpc = (TextPropCollection)
				_styleAtom.GetParagraphStyles().GetLast();
			tpc.updateTextSize(
					tpc.GetCharactersCovered() - pOverRun
			);
		}
		if(cOverRun > 0) {
			TextPropCollection tpc = (TextPropCollection)
				_styleAtom.GetCharacterStyles().GetLast();
			tpc.updateTextSize(
					tpc.GetCharactersCovered() - cOverRun
			);
		}

		// Next, add the styles for its paragraph and characters
		TextPropCollection newPTP =
			_styleAtom.AddParagraphTextPropCollection(s.Length+pOverRun);
		TextPropCollection newCTP =
			_styleAtom.AddCharacterTextPropCollection(s.Length+cOverRun);

		// Now, create the new RichTextRun
		RichTextRun nr = new RichTextRun(
				this, oldSize, s.Length,
				newPTP, newCTP, false, false
		);

		// Add the new RichTextRun onto our list
		RichTextRun[] newRuns = new RichTextRun[_rtRuns.Length+1];
		Array.Copy(_rtRuns, 0, newRuns, 0, _rtRuns.Length);
		newRuns[newRuns.Length-1] = nr;
		_rtRuns = newRuns;

		// And return the new run to the caller
		return nr;
	}

	/**
	 * Saves the given string to the records. Doesn't
	 *  touch the stylings.
	 */
	private void storeText(String s) {
		// Remove a single trailing \r, as there is an implicit one at the
		//  end of every record
		if(s\.EndsWith("\r")) {
			s = s.Substring(0, s.Length-1);
		}

		// Store in the appropriate record
		if(_isUnicode) {
			// The atom can safely convert to unicode
			_charAtom.SetText(s);
		} else {
			// Will it fit in a 8 bit atom?
			bool HasMultibyte = StringUtil.HasMultibyte(s);
			if(! hasMultibyte) {
				// Fine to go into 8 bit atom
				byte[] text = new byte[s.Length];
				StringUtil.PutCompressedUnicode(s,text,0);
				_byteAtom.SetText(text);
			} else {
				// Need to swap a TextBytesAtom for a TextCharsAtom

				// Build the new TextCharsAtom
				_charAtom = new TextCharsAtom();
				_charAtom.SetText(s);

				// Use the TextHeaderAtom to do the swap on the parent
				RecordContainer parent = _headerAtom.GetParentRecord();
				Record[] cr = parent.GetChildRecords();
				for(int i=0; i<cr.Length; i++) {
					// Look for TextBytesAtom
					if(cr[i].Equals(_byteAtom)) {
						// Found it, so Replace, then all done
						cr[i] = _charAtom;
						break;
					}
				}

				// Flag the change
				_byteAtom = null;
				_isUnicode = true;
			}
		}
        /**
         * If TextSpecInfoAtom is present, we must update the text size in it,
         * otherwise the ppt will be corrupted
         */
        if(_records != null) for (int i = 0; i < _records.Length; i++) {
            if(_records[i] is TextSpecInfoAtom){
                TextSpecInfoAtom specAtom = (TextSpecInfoAtom)_records[i];
                if((s.Length + 1) != specAtom.GetCharactersCovered()){
                    specAtom.reset(s.Length + 1);
                }
            }
        }
	}

	/**
	 * Handles an update to the text stored in one of the Rich Text Runs
	 * @param run
	 * @param s
	 */
	public void ChangeTextInRichTextRun(RichTextRun Run, String s) {
		// Figure out which run it is
		int RunID = -1;
		for(int i=0; i<_rtRuns.Length; i++) {
			if(Run.Equals(_rtRuns[i])) {
				RunID = i;
			}
		}
		if(RunID == -1) {
			throw new ArgumentException("Supplied RichTextRun wasn't from this TextRun");
		}

		// Ensure a StyleTextPropAtom is present, Adding if required
		ensureStyleAtomPresent();

		// Update the text length for its Paragraph and Character stylings
		// If it's shared:
		//   * calculate the new length based on the Run's old text
		//   * this should leave in any +1's for the end of block if needed
		// If it isn't shared:
		//   * reset the length, to the new string's length
		//   * add on +1 if the last block
		// The last run needs its stylings to be 1 longer than the raw
		//  text is. This is to define the stylings that any new text
		//  that is Added will inherit
		TextPropCollection pCol = Run._getRawParagraphStyle();
		TextPropCollection cCol = Run._getRawCharacterStyle();
		int newSize = s.Length;
		if(RunID == _rtRuns.Length-1) {
			newSize++;
		}

		if(Run._isParagraphStyleShared()) {
			pCol.updateTextSize( pCol.GetCharactersCovered() - Run.GetLength() + s.Length );
		} else {
			pCol.updateTextSize(newSize);
		}
		if(Run._isCharacterStyleShared()) {
			cCol.updateTextSize( cCol.GetCharactersCovered() - Run.GetLength() + s.Length );
		} else {
			cCol.updateTextSize(newSize);
		}

		// Build up the new text
		// As we go through, update the start position for all subsequent Runs
		// The building relies on the old text still being present
		StringBuilder newText = new StringBuilder();
		for(int i=0; i<_rtRuns.Length; i++) {
			int newStartPos = newText.Length;

			// Build up the new text
			if(i != RunID) {
				// Not the affected Run, so keep old text
				newText.Append(_rtRuns[i].GetRawText());
			} else {
				// Affected Run, so use new text
				newText.Append(s);
			}

			// Do we need to update the start position of this Run?
			// (Need to Get the text before we update the start pos)
			if(i <= RunID) {
				// Change is after this, so don't need to change start position
			} else {
				// Change has occured, so update start position
				_rtRuns[i].updateStartPosition(newStartPos);
			}
		}

		// Now we can save the new text
		storeText(newText.ToString());
	}

	/**
	 * Changes the text, and Sets it all to have the same styling
	 *  as the the first character has.
	 * If you care about styling, do SetText on a RichTextRun instead
	 */
	public void SetRawText(String s) {
		// Save the new text to the atoms
		storeText(s);
		RichTextRun fst = _rtRuns[0];

		// Finally, zap and re-do the RichTextRuns
		for(int i=0; i<_rtRuns.Length; i++) { _rtRuns[i] = null; }
		_rtRuns = new RichTextRun[1];
        _rtRuns[0] = fst;

		// Now handle record stylings:
		// If there isn't styling
		//  no Change, stays with no styling
		// If there is styling:
		//  everthing Gets the same style that the first block has
		if(_styleAtom != null) {
			LinkedList pStyles = _styleAtom.GetParagraphStyles();
			while(pStyles.Count > 1) { pStyles.RemoveLast(); }

			LinkedList cStyles = _styleAtom.GetCharacterStyles();
			while(cStyles.Count > 1) { cStyles.RemoveLast(); }

			_rtRuns[0].SetText(s);
		} else {
			// Recreate rich text run with no styling
			_rtRuns[0] = new RichTextRun(this,0,s.Length);
		}

	}

    /**
     * Changes the text.
     * Converts '\r' into '\n'
     */
    public void SetText(String s) {
        String text = normalize(s);
        SetRawText(text);
    }

    /**
	 * Ensure a StyleTextPropAtom is present for this Run,
	 *  by Adding if required. Normally for internal TextRun use.
	 */
	public void ensureStyleAtomPresent() {
		if(_styleAtom != null) {
			// All there
			return;
		}

		// Create a new one at the right size
		_styleAtom = new StyleTextPropAtom(getRawText().Length + 1);

		// Use the TextHeader atom to Get at the parent
		RecordContainer RunAtomsParent = _headerAtom.GetParentRecord();

		// Add the new StyleTextPropAtom after the TextCharsAtom / TextBytesAtom
		Record AddAfter = _byteAtom;
		if(_byteAtom == null) { AddAfter = _charAtom; }
		RunAtomsParent.AddChildAfter(_styleAtom, AddAfter);

		// Feed this to our sole rich text run
		if(_rtRuns.Length != 1) {
			throw new InvalidOperationException("Needed to add StyleTextPropAtom when had many rich text Runs");
		}
		// These are the only styles for now
		_rtRuns[0].supplyTextProps(
				(TextPropCollection)_styleAtom.GetParagraphStyles().Get(0),
				(TextPropCollection)_styleAtom.GetCharacterStyles().Get(0),
				false,
				false
		);
	}

	// Accesser methods follow

	/**
	 * Returns the text content of the Run, which has been made safe
	 * for printing and other use.
	 */
	public String GetText() {
		String rawText = GetRawText();

		// PowerPoint seems to store files with \r as the line break
		// The messes things up on everything but a Mac, so translate
		//  them to \n
		String text = rawText.Replace('\r','\n');

        int type = _headerAtom == null ? 0 : _headerAtom.GetTextType();
        if(type == TextHeaderAtom.TITLE_TYPE || type == TextHeaderAtom.CENTER_TITLE_TYPE){
            //0xB acts like cariage return in page titles and like blank in the others
            text = text.Replace((char) 0x0B, '\n');
        } else {
            text = text.Replace((char) 0x0B, ' ');
        }
		return text;
	}

	/**
	* Returns the raw text content of the Run. This hasn't had any
	*  Changes applied to it, and so is probably unlikely to print
	*  out nicely.
	*/
	public String GetRawText() {
		if(_isUnicode) {
			return _charAtom.GetText();
		}
		return _byteAtom.GetText();
	}

	/**
	 * Fetch the rich text Runs (Runs of text with the same styling) that
	 *  are Contained within this block of text
	 */
	public RichTextRun[] GetRichTextRuns() {
		return 	_rtRuns;
	}

	/**
	* Returns the type of the text, from the TextHeaderAtom.
	* Possible values can be seen from TextHeaderAtom
	* @see NPOI.HSLF.record.TextHeaderAtom
	*/
	public int GetRunType() {
		return _headerAtom.GetTextType();
	}

	/**
	* Changes the type of the text. Values should be taken
	*  from TextHeaderAtom. No Checking is done to ensure you
	*  Set this to a valid value!
	* @see NPOI.HSLF.record.TextHeaderAtom
	*/
	public void SetRunType(int type) {
		_headerAtom.SetTextType(type);
	}

	/**
	 * Supply the SlideShow we belong to.
	 * Also passes it on to our child RichTextRuns
	 */
	public void supplySlideShow(SlideShow ss) {
		slideShow = ss;
		if(_rtRuns != null) {
			for(int i=0; i<_rtRuns.Length; i++) {
				_rtRuns[i].supplySlideShow(slideShow);
			}
		}
	}

    public void SetSheet(Sheet sheet){
        this._sheet = sheet;
    }

    public Sheet Sheet{
        return this._sheet;
    }

    /**
     * @return  Shape ID
     */
    protected int GetShapeId(){
        return shapeId;
    }

    /**
     *  @param id Shape ID
     */
    protected void SetShapeId(int id){
        shapeId = id;
    }

    /**
     * @return  0-based index of the text run in the SLWT Container
     */
    protected int GetIndex(){
        return slwtIndex;
    }

    /**
     *  @param id 0-based index of the text run in the SLWT Container
     */
    protected void SetIndex(int id){
        slwtIndex = id;
    }

    /**
     * Returns the array of all hyperlinks in this text run
     *
     * @return the array of all hyperlinks in this text run
     * or <code>null</code> if not found.
     */
    public Hyperlink[] GetHyperlinks(){
        return Hyperlink.Find(this);
    }

    /**
     * Fetch RichTextRun at a given position
     *
     * @param pos 0-based index in the text
     * @return RichTextRun or null if not found
     */
    public RichTextRun GetRichTextRunAt(int pos){
        for (int i = 0; i < _rtRuns.Length; i++) {
            int start = _rtRuns[i].GetStartIndex();
            int end = _rtRuns[i].GetEndIndex();
            if(pos >= start && pos < end) return _rtRuns[i];
        }
        return null;
    }

    public TextRulerAtom GetTextRuler(){
        if(_ruler == null){
            if(_records != null) for (int i = 0; i < _records.Length; i++) {
                if(_records[i] is TextRulerAtom) {
                    _ruler = (TextRulerAtom)_records[i];
                    break;
                }
            }

        }
        return _ruler;

    }

    public TextRulerAtom CreateTextRuler(){
        _ruler = GetTextRuler();
        if(_ruler == null){
            _ruler = TextRulerAtom.GetParagraphInstance();
            _headerAtom.GetParentRecord().AppendChildRecord(_ruler);
        }
        return _ruler;
    }

    /**
     * Returns a new string with line breaks Converted into internal ppt representation
     */
    public String normalize(String s){
        String ns = s.ReplaceAll("\\r?\\n", "\r");
        return ns;
    }

    /**
     * Returns records that make up this text run
     *
     * @return text run records
     */
    public Record[] GetRecords(){
        return _records;
    }

}





