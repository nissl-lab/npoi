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
using NPOI.ddf.EscherOptRecord;
using NPOI.ddf.EscherProperties;
using NPOI.ddf.EscherSimpleProperty;
using NPOI.ddf.EscherSpRecord;
using NPOI.ddf.EscherTextboxRecord;
using NPOI.HSLF.exceptions.HSLFException;
using NPOI.HSLF.record.EscherTextboxWrapper;
using NPOI.HSLF.record.InteractiveInfo;
using NPOI.HSLF.record.InteractiveInfoAtom;
using NPOI.HSLF.record.OEPlaceholderAtom;
using NPOI.HSLF.record.OutlineTextRefAtom;
using NPOI.HSLF.record.PPDrawing;
using NPOI.HSLF.record.Record;
using NPOI.HSLF.record.RecordTypes;
using NPOI.HSLF.record.StyleTextPropAtom;
using NPOI.HSLF.record.TextCharsAtom;
using NPOI.HSLF.record.TextHeaderAtom;
using NPOI.HSLF.record.TxInteractiveInfoAtom;
using NPOI.HSLF.usermodel.RichTextRun;
using NPOI.util.POILogger;

/**
 * A common superclass of all shapes that can hold text.
 *
 * @author Yegor Kozlov
 */
public abstract class TextShape : SimpleShape {

    /**
     * How to anchor the text
     */
    public static int AnchorTop = 0;
    public static int AnchorMiddle = 1;
    public static int AnchorBottom = 2;
    public static int AnchorTopCentered = 3;
    public static int AnchorMiddleCentered = 4;
    public static int AnchorBottomCentered = 5;
    public static int AnchorTopBaseline = 6;
    public static int AnchorBottomBaseline = 7;
    public static int AnchorTopCenteredBaseline = 8;
    public static int AnchorBottomCenteredBaseline = 9;

    /**
     * How to wrap the text
     */
    public static int WrapSquare = 0;
    public static int WrapByPoints = 1;
    public static int WrapNone = 2;
    public static int WrapTopBottom = 3;
    public static int WrapThrough = 4;

    /**
     * How to align the text
     */
    public static int AlignLeft = 0;
    public static int AlignCenter = 1;
    public static int AlignRight = 2;
    public static int AlignJustify = 3;

    /**
     * TextRun object which holds actual text and format data
     */
    protected TextRun _txtRun;

    /**
     * Escher Container which holds text attributes such as
     * TextHeaderAtom, TextBytesAtom ot TextCharsAtom, StyleTextPropAtom etc.
     */
    protected EscherTextboxWrapper _txtbox;

    /**
     * Used to calculate text bounds
     */
    protected static FontRenderContext _frc = new FontRenderContext(null, true, true);

    /**
     * Create a TextBox object and Initialize it from the supplied Record Container.
     *
     * @param escherRecord       <code>EscherSpContainer</code> Container which holds information about this shape
     * @param parent    the parent of the shape
     */
   protected TextShape(EscherContainerRecord escherRecord, Shape parent){
        base(escherRecord, parent);

    }

    /**
     * Create a new TextBox. This constructor is used when a new shape is Created.
     *
     * @param parent    the parent of this Shape. For example, if this text box is a cell
     * in a table then the parent is Table.
     */
    public TextShape(Shape parent){
        base(null, parent);
        _escherContainer = CreateSpContainer(parent is ShapeGroup);
    }

    /**
     * Create a new TextBox. This constructor is used when a new shape is Created.
     *
     */
    public TextShape(){
        this(null);
    }

    public TextRun CreateTextRun(){
        _txtbox = GetEscherTextboxWrapper();
        if(_txtbox == null) _txtbox = new EscherTextboxWrapper();

        _txtrun = GetTextRun();
        if(_txtrun == null){
            TextHeaderAtom tha = new TextHeaderAtom();
            tha.SetParentRecord(_txtbox);
            _txtbox.AppendChildRecord(tha);

            TextCharsAtom tca = new TextCharsAtom();
            _txtbox.AppendChildRecord(tca);

            StyleTextPropAtom sta = new StyleTextPropAtom(0);
            _txtbox.AppendChildRecord(sta);

            _txtrun = new TextRun(tha,tca,sta);
            _txtRun._records = new Record[]{tha, tca, sta};
            _txtRun.SetText("");

            _escherContainer.AddChildRecord(_txtbox.GetEscherRecord());

            SetDefaultTextProperties(_txtRun);
        }

        return _txtRun;
    }

    /**
     * Set default properties for the  TextRun.
     * Depending on the text and shape type the defaults are different:
     *   TextBox: align=left, valign=top
     *   AutoShape: align=center, valign=middle
     *
     */
    protected void SetDefaultTextProperties(TextRun _txtRun){

    }

    /**
     * Returns the text Contained in this text frame.
     *
     * @return the text string for this textbox.
     */
     public String GetText(){
        TextRun tx = GetTextRun();
        return tx == null ? null : tx.GetText();
    }

    /**
     * Sets the text Contained in this text frame.
     *
     * @param text the text string used by this object.
     */
    public void SetText(String text){
        TextRun tx = GetTextRun();
        if(tx == null){
            tx = CreateTextRun();
        }
        tx.SetText(text);
        SetTextId(text.HashCode());
    }

    /**
     * When a textbox is Added to  a sheet we need to tell upper-level
     * <code>PPDrawing</code> about it.
     *
     * @param sh the sheet we are Adding to
     */
    protected void afterInsert(Sheet sh){
        super.afterInsert(sh);

        EscherTextboxWrapper _txtbox = GetEscherTextboxWrapper();
        if(_txtbox != null){
            PPDrawing ppdrawing = sh.GetPPDrawing();
            ppdrawing.AddTextboxWrapper(_txtbox);
            // Ensure the escher layer knows about the Added records
            try {
                _txtbox.WriteOut(null);
            } catch (IOException e){
                throw new HSLFException(e);
            }
            if(getAnchor().Equals(new Rectangle()) && !"".Equals(getText())) resizeToFitText();
        }
        if(_txtrun != null) {
            _txtRun.SetShapeId(getShapeId());
            sh.onAddTextShape(this);
        }
    }

    protected EscherTextboxWrapper GetEscherTextboxWrapper(){
        if(_txtbox == null){
            EscherTextboxRecord textRecord = (EscherTextboxRecord)Shape.GetEscherChild(_escherContainer, EscherTextboxRecord.RECORD_ID);
            if(textRecord != null) _txtbox = new EscherTextboxWrapper(textRecord);
        }
        return _txtbox;
    }
    /**
     * Adjust the size of the TextShape so it encompasses the text inside it.
     *
     * @return a <code>Rectangle2D</code> that is the bounds of this <code>TextShape</code>.
     */
    public Rectangle2D resizeToFitText(){
        String txt = GetText();
        if(txt == null || txt.Length == 0) return new Rectangle2D.Float();

        RichTextRun rt = GetTextRun().GetRichTextRuns()[0];
        int size = rt.GetFontSize();
        int style = 0;
        if (rt.IsBold()) style |= Font.BOLD;
        if (rt.IsItalic()) style |= Font.ITALIC;
        String fntname = rt.GetFontName();
        Font font = new Font(fntname, style, size);

        float width = 0, height = 0, leading = 0;
        String[] lines = txt.split("\n");
        for (int i = 0; i < lines.Length; i++) {
            if(lines[i].Length == 0) continue;

            TextLayout layout = new TextLayout(lines[i], font, _frc);

            leading = Math.max(leading, layout.GetLeading());
            width = Math.max(width, layout.GetAdvance());
            height = Math.max(height, (height + (layout.GetDescent() + layout.GetAscent())));
        }

        // add one character to width
        Rectangle2D charBounds = font.GetMaxCharBounds(_frc);
        width += GetMarginLeft() + GetMarginRight() + charBounds.Width;

        // add leading to height
        height += GetMarginTop() + GetMarginBottom() + leading;

        Rectangle2D anchor = GetAnchor2D();
        anchor.SetRect(anchor.GetX(), anchor.GetY(), width, height);
        SetAnchor(anchor);

        return anchor;
    }

    /**
     * Returns the type of vertical alignment for the text.
     * One of the <code>Anchor*</code> constants defined in this class.
     *
     * @return the type of alignment
     */
    public int GetVerticalAlignment(){
        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
        EscherSimpleProperty prop = (EscherSimpleProperty)getEscherProperty(opt, EscherProperties.TEXT__ANCHORTEXT);
        int valign = TextShape.AnchorTop;
        if (prop == null){
            /**
             * If vertical alignment was not found in the shape properties then try to
             * fetch the master shape and search for the align property there.
             */
            int type = GetTextRun().GetRunType();
            MasterSheet master = Sheet.GetMasterSheet();
            if(master != null){
                TextShape masterShape = master.GetPlaceholderByTextType(type);
                if(masterShape != null) valign = masterShape.GetVerticalAlignment();
            } else {
                //not found in the master sheet. Use the hardcoded defaults.
                switch (type){
                     case TextHeaderAtom.TITLE_TYPE:
                     case TextHeaderAtom.CENTER_TITLE_TYPE:
                         valign = TextShape.AnchorMiddle;
                         break;
                     default:
                         valign = TextShape.AnchorTop;
                         break;
                 }
            }
        } else {
            valign = prop.GetPropertyValue();
        }
        return valign;
    }

    /**
     * Sets the type of vertical alignment for the text.
     * One of the <code>Anchor*</code> constants defined in this class.
     *
     * @param align - the type of alignment
     */
    public void SetVerticalAlignment(int align){
        SetEscherProperty(EscherProperties.TEXT__ANCHORTEXT, align);
    }

    /**
     * Sets the type of horizontal alignment for the text.
     * One of the <code>Align*</code> constants defined in this class.
     *
     * @param align - the type of horizontal alignment
     */
    public void SetHorizontalAlignment(int align){
        TextRun tx = GetTextRun();
        if(tx != null) tx.GetRichTextRuns()[0].SetAlignment(align);
    }

    /**
     * Gets the type of horizontal alignment for the text.
     * One of the <code>Align*</code> constants defined in this class.
     *
     * @return align - the type of horizontal alignment
     */
    public int GetHorizontalAlignment(){
        TextRun tx = GetTextRun();
        return tx == null ? -1 : tx.GetRichTextRuns()[0].GetAlignment();
    }

    /**
     * Returns the distance (in points) between the bottom of the text frame
     * and the bottom of the inscribed rectangle of the shape that Contains the text.
     * Default value is 1/20 inch.
     *
     * @return the botom margin
     */
    public float GetMarginBottom(){
        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
        EscherSimpleProperty prop = (EscherSimpleProperty)getEscherProperty(opt, EscherProperties.TEXT__TEXTBOTTOM);
        int val = prop == null ? EMU_PER_INCH/20 : prop.GetPropertyValue();
        return (float)val/EMU_PER_POINT;
    }

    /**
     * Sets the botom margin.
     * @see #getMarginBottom()
     *
     * @param margin    the bottom margin
     */
    public void SetMarginBottom(float margin){
        SetEscherProperty(EscherProperties.TEXT__TEXTBOTTOM, (int)(margin*EMU_PER_POINT));
    }

    /**
     *  Returns the distance (in points) between the left edge of the text frame
     *  and the left edge of the inscribed rectangle of the shape that Contains
     *  the text.
     *  Default value is 1/10 inch.
     *
     * @return the left margin
     */
    public float GetMarginLeft(){
        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
        EscherSimpleProperty prop = (EscherSimpleProperty)getEscherProperty(opt, EscherProperties.TEXT__TEXTLEFT);
        int val = prop == null ? EMU_PER_INCH/10 : prop.GetPropertyValue();
        return (float)val/EMU_PER_POINT;
    }

    /**
     * Sets the left margin.
     * @see #getMarginLeft()
     *
     * @param margin    the left margin
     */
    public void SetMarginLeft(float margin){
        SetEscherProperty(EscherProperties.TEXT__TEXTLEFT, (int)(margin*EMU_PER_POINT));
    }

    /**
     *  Returns the distance (in points) between the right edge of the
     *  text frame and the right edge of the inscribed rectangle of the shape
     *  that Contains the text.
     *  Default value is 1/10 inch.
     *
     * @return the right margin
     */
    public float GetMarginRight(){
        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
        EscherSimpleProperty prop = (EscherSimpleProperty)getEscherProperty(opt, EscherProperties.TEXT__TEXTRIGHT);
        int val = prop == null ? EMU_PER_INCH/10 : prop.GetPropertyValue();
        return (float)val/EMU_PER_POINT;
    }

    /**
     * Sets the right margin.
     * @see #getMarginRight()
     *
     * @param margin    the right margin
     */
    public void SetMarginRight(float margin){
        SetEscherProperty(EscherProperties.TEXT__TEXTRIGHT, (int)(margin*EMU_PER_POINT));
    }

     /**
     *  Returns the distance (in points) between the top of the text frame
     *  and the top of the inscribed rectangle of the shape that Contains the text.
     *  Default value is 1/20 inch.
     *
     * @return the top margin
     */
    public float GetMarginTop(){
        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
        EscherSimpleProperty prop = (EscherSimpleProperty)getEscherProperty(opt, EscherProperties.TEXT__TEXTTOP);
        int val = prop == null ? EMU_PER_INCH/20 : prop.GetPropertyValue();
        return (float)val/EMU_PER_POINT;
    }

   /**
     * Sets the top margin.
     * @see #getMarginTop()
     *
     * @param margin    the top margin
     */
    public void SetMarginTop(float margin){
        SetEscherProperty(EscherProperties.TEXT__TEXTTOP, (int)(margin*EMU_PER_POINT));
    }


    /**
     * Returns the value indicating word wrap.
     *
     * @return the value indicating word wrap.
     *  Must be one of the <code>Wrap*</code> constants defined in this class.
     */
    public int GetWordWrap(){
        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
        EscherSimpleProperty prop = (EscherSimpleProperty)getEscherProperty(opt, EscherProperties.TEXT__WRAPTEXT);
        return prop == null ? WrapSquare : prop.GetPropertyValue();
    }

    /**
     *  Specifies how the text should be wrapped
     *
     * @param wrap  the value indicating how the text should be wrapped.
     *  Must be one of the <code>Wrap*</code> constants defined in this class.
     */
    public void SetWordWrap(int wrap){
        SetEscherProperty(EscherProperties.TEXT__WRAPTEXT, wrap);
    }

    /**
     * @return id for the text.
     */
    public int GetTextId(){
        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
        EscherSimpleProperty prop = (EscherSimpleProperty)getEscherProperty(opt, EscherProperties.TEXT__TEXTID);
        return prop == null ? 0 : prop.GetPropertyValue();
    }

    /**
     * Sets text ID
     *
     * @param id of the text
     */
    public void SetTextId(int id){
        SetEscherProperty(EscherProperties.TEXT__TEXTID, id);
    }

    /**
      * @return the TextRun object for this text box
      */
     public TextRun GetTextRun(){
         if(_txtrun == null) InitTextRun();
         return _txtRun;
     }

    public void SetSheet(Sheet sheet) {
        _sheet = sheet;

        // Initialize _txtrun object.
        // (We can't do it in the constructor because the sheet
        //  is not assigned then, it's only built once we have
        //  all the records)
        TextRun tx = GetTextRun();
        if (tx != null) {
            // Supply the sheet to our child RichTextRuns
            tx.SetSheet(_sheet);
            RichTextRun[] rt = tx.GetRichTextRuns();
            for (int i = 0; i < rt.Length; i++) {
                rt[i].supplySlideShow(_sheet.GetSlideShow());
            }
        }

    }

    protected void InitTextRun(){
        EscherTextboxWrapper txtbox = GetEscherTextboxWrapper();
        Sheet sheet = Sheet;

        if(sheet == null || txtbox == null) return;

        OutlineTextRefAtom ota = null;

        Record[] child = txtbox.GetChildRecords();
        for (int i = 0; i < child.Length; i++) {
            if (child[i] is OutlineTextRefAtom) {
                ota = (OutlineTextRefAtom)child[i];
                break;
            }
        }

        TextRun[] Runs = _sheet.GetTextRuns();
        if (ota != null) {
            int idx = ota.GetTextIndex();
            for (int i = 0; i < Runs.Length; i++) {
                if(Runs[i].GetIndex() == idx){
                    _txtrun = Runs[i];
                    break;
                }
            }
            if(_txtrun == null) {
                logger.log(POILogger.WARN, "text run not found for OutlineTextRefAtom.TextIndex=" + idx);
            }
        } else {
            EscherSpRecord escherSpRecord = _escherContainer.GetChildById(EscherSpRecord.RECORD_ID);
            int shapeId = escherSpRecord.GetShapeId();
            if(Runs != null) for (int i = 0; i < Runs.Length; i++) {
                if(Runs[i].GetShapeId() == shapeId){
                    _txtrun = Runs[i];
                    break;
                }
            }
        }
        // ensure the same references child records of TextRun
        if(_txtrun != null) for (int i = 0; i < child.Length; i++) {
            foreach (Record r in _txtRun.GetRecords()) {
                if (child[i].GetRecordType() == r.GetRecordType()) {
                    child[i] = r;
                }
            }
        }
    }

    public void Draw(Graphics2D graphics){
        AffineTransform at = graphics.GetTransform();
        ShapePainter.paint(this, graphics);
        new TextPainter(this).paint(graphics);
        graphics.SetTransform(at);
    }

    /**
     * Return <code>OEPlaceholderAtom</code>, the atom that describes a placeholder.
     *
     * @return <code>OEPlaceholderAtom</code> or <code>null</code> if not found
     */
    public OEPlaceholderAtom GetPlaceholderAtom(){
        return (OEPlaceholderAtom)getClientDataRecord(RecordTypes.OEPlaceholderAtom.typeID);
    }

    /**
     *
     * Assigns a hyperlink to this text shape
     *
     * @param linkId    id of the hyperlink, @see NPOI.HSLF.usermodel.SlideShow#AddHyperlink(Hyperlink)
     * @param      beginIndex   the beginning index, inclusive.
     * @param      endIndex     the ending index, exclusive.
     * @see NPOI.HSLF.usermodel.SlideShow#AddHyperlink(Hyperlink)
     */
    public void SetHyperlink(int linkId, int beginIndex, int endIndex){
        //TODO validate beginIndex and endIndex and throw ArgumentException

        InteractiveInfo info = new InteractiveInfo();
        InteractiveInfoAtom infoAtom = info.GetInteractiveInfoAtom();
        infoAtom.SetAction(InteractiveInfoAtom.ACTION_HYPERLINK);
        infoAtom.SetHyperlinkType(InteractiveInfoAtom.LINK_Url);
        infoAtom.SetHyperlinkID(linkId);

        _txtbox.AppendChildRecord(info);

        TxInteractiveInfoAtom txiatom = new TxInteractiveInfoAtom();
        txiatom.SetStartIndex(beginIndex);
        txiatom.SetEndIndex(endIndex);
        _txtbox.AppendChildRecord(txiatom);

    }

}





