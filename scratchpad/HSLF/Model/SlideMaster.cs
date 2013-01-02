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

using NPOI.HSLF.Model.textproperties.TextProp;
using NPOI.HSLF.Model.textproperties.TextPropCollection;
using NPOI.HSLF.record.*;
using NPOI.HSLF.usermodel.SlideShow;

/**
 * SlideMaster determines the graphics, layout, and formatting for all the slides in a given presentation.
 * It stores information about default font styles, placeholder sizes and positions,
 * background design, and color schemes.
 *
 * @author Yegor Kozlov
 */
public class SlideMaster : MasterSheet {
    private TextRun[] _Runs;

    /**
     * all TxMasterStyleAtoms available in this master
     */
    private TxMasterStyleAtom[] _txmaster;

    /**
     * Constructs a SlideMaster from the MainMaster record,
     *
     */
    public SlideMaster(MainMaster record, int sheetNo) {
        base(record, sheetNo);

        _Runs = FindTextRuns(getPPDrawing());
        for (int i = 0; i < _Runs.Length; i++) _Runs[i].SetSheet(this);
    }

    /**
     * Returns an array of all the TextRuns found
     */
    public TextRun[] GetTextRuns() {
        return _Runs;
    }

    /**
     * Returns <code>null</code> since SlideMasters doen't have master sheet.
     */
    public MasterSheet GetMasterSheet() {
        return null;
    }

    /**
     * Pickup a style attribute from the master.
     * This is the "workhorse" which returns the default style attrubutes.
     */
    public TextProp GetStyleAttribute(int txtype, int level, String name, bool IsCharacter) {

        TextProp prop = null;
        for (int i = level; i >= 0; i--) {
            TextPropCollection[] styles =
                    isCharacter ? _txmaster[txtype].GetCharacterStyles() : _txmaster[txtype].GetParagraphStyles();
            if (i < styles.Length) prop = styles[i].FindByName(name);
            if (prop != null) break;
        }
        if (prop == null) {
            if(isCharacter) {
                switch (txtype) {
                    case TextHeaderAtom.CENTRE_BODY_TYPE:
                    case TextHeaderAtom.HALF_BODY_TYPE:
                    case TextHeaderAtom.QUARTER_BODY_TYPE:
                        txtype = TextHeaderAtom.BODY_TYPE;
                        break;
                    case TextHeaderAtom.CENTER_TITLE_TYPE:
                        txtype = TextHeaderAtom.TITLE_TYPE;
                        break;
                    default:
                        return null;
                }
            } else {
                switch (txtype) {
                    case TextHeaderAtom.CENTRE_BODY_TYPE:
                    case TextHeaderAtom.HALF_BODY_TYPE:
                    case TextHeaderAtom.QUARTER_BODY_TYPE:
                        txtype = TextHeaderAtom.BODY_TYPE;
                        break;
                    case TextHeaderAtom.CENTER_TITLE_TYPE:
                        txtype = TextHeaderAtom.TITLE_TYPE;
                        break;
                    default:
                        return null;
                }
            }
            prop = GetStyleAttribute(txtype, level, name, isCharacter);
        }
        return prop;
    }

    /**
     * Assign SlideShow for this slide master.
     * (Used interanlly)
     */
    public void SetSlideShow(SlideShow ss) {
        super.SetSlideShow(ss);

        //after the slide show is assigned collect all available style records
        if (_txmaster == null) {
            _txmaster = new TxMasterStyleAtom[9];

            TxMasterStyleAtom txdoc = GetSlideShow().GetDocumentRecord().GetEnvironment().GetTxMasterStyleAtom();
            _txmaster[txdoc.GetTextType()] = txdoc;

            TxMasterStyleAtom[] txrec = ((MainMaster)getSheetContainer()).GetTxMasterStyleAtoms();
            for (int i = 0; i < txrec.Length; i++) {
                _txmaster[txrec[i].GetTextType()] = txrec[i];
            }
        }
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

    public TxMasterStyleAtom[] GetTxMasterStyleAtoms(){
        return _txmaster;
    }
}





