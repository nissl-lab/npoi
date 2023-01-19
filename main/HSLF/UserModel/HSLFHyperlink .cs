/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using NPOI.Common.UserModel;
using NPOI.HSLF.Record;
using NPOI.SL.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.HSLF.UserModel
{
	public class HSLFHyperlink: Hyperlink<HSLFShape, HSLFTextParagraph>
	{
		private ExHyperlink exHyper;
    private InteractiveInfo info;
    private TxInteractiveInfoAtom txinfo;

    protected HSLFHyperlink(ExHyperlink exHyper, InteractiveInfo info) {
        this.info = info;
        this.exHyper = exHyper;
    }

    public ExHyperlink GetExHyperlink() {
        return exHyper;
    }

    public InteractiveInfo GetInfo() {
        return info;
    }

    public TxInteractiveInfoAtom GetTextRunInfo() {
        return txinfo;
    }

    protected void SetTextRunInfo(TxInteractiveInfoAtom txinfo) {
        this.txinfo = txinfo;
    }

    /**
     * Creates a new Hyperlink and assign it to a shape
     * This is only a helper method - use {@link HSLFSimpleShape#createHyperlink()} instead!
     *
     * @param shape the shape which receives the hyperlink
     * @return the new hyperlink
     *
     * @see HSLFSimpleShape#createHyperlink()
     */
    /* package */ static HSLFHyperlink CreateHyperlink(HSLFSimpleShape shape) {
        // TODO: check if a hyperlink already exists
        ExHyperlink exHyper = new ExHyperlink();
            ///TODO: Add CreateHyperlink functionality
        //int linkId = shape.GetSheet().GetSlideShow().AddToObjListAtom(exHyper);
        //ExHyperlinkAtom obj = exHyper.GetExHyperlinkAtom();
        //obj.SetNumber(linkId);
        InteractiveInfo info = new InteractiveInfo();
        //info.GetInteractiveInfoAtom().SetHyperlinkID(linkId);
        //HSLFEscherClientDataRecord cldata = shape.GetClientData(true);
        //cldata.AddChild(info);
        HSLFHyperlink hyper = new HSLFHyperlink(exHyper, info);
        //hyper.linkToNextSlide();
        //shape.SetHyperlink(hyper);
        return hyper;
    }

    /**
     * Creates a new Hyperlink for a textrun.
     * This is only a helper method - use {@link HSLFTextRun#createHyperlink()} instead!
     *
     * @param run the run which receives the hyperlink
     * @return the new hyperlink
     *
     * @see HSLFTextRun#createHyperlink()
     */
    /* package */ static HSLFHyperlink CreateHyperlink(HSLFTextRun run) {
        // TODO: check if a hyperlink already exists
        ExHyperlink exHyper = new ExHyperlink();
        //int linkId = run.GetTextParagraph().GetSheet().GetSlideShow().AddToObjListAtom(exHyper);
        ExHyperlinkAtom obj = exHyper.GetExHyperlinkAtom();
        //obj.SetNumber(linkId);
        InteractiveInfo info = new InteractiveInfo();
        //info.GetInteractiveInfoAtom().SetHyperlinkID(linkId);
        // don't add the hyperlink now to text paragraph records
        // this will be done, when the paragraph is saved
        HSLFHyperlink hyper = new HSLFHyperlink(exHyper, info);
        //hyper.linkToNextSlide();

        TxInteractiveInfoAtom txinfo = new TxInteractiveInfoAtom();
        //int startIdx = run.GetTextParagraph().GetStartIdxOfTextRun(run);
        //int endIdx = startIdx + run.GetLength();
        //txinfo.SetStartIndex(startIdx);
        //txinfo.SetEndIndex(endIdx);
        hyper.SetTextRunInfo(txinfo);

        //run.SetHyperlink(hyper);
        return hyper;
    }


    /**
     * Gets the type of the hyperlink action.
     * Must be a {@code LINK_*} constant
     *
     * @return the hyperlink URL
     * @see InteractiveInfoAtom
     */
    //@Override
    public HyperlinkType GetType() {
        switch (info.GetInteractiveInfoAtom().GetHyperlinkType()) {
            case InteractiveInfoAtom.LINK_Url:
                return (exHyper.GetLinkURL().StartsWith("mailto:")) ? HyperlinkType.EMAIL : HyperlinkType.URL;
            case InteractiveInfoAtom.LINK_NextSlide:
            case InteractiveInfoAtom.LINK_PreviousSlide:
            case InteractiveInfoAtom.LINK_FirstSlide:
            case InteractiveInfoAtom.LINK_LastSlide:
            case InteractiveInfoAtom.LINK_SlideNumber:
                return HyperlinkType.DOCUMENT;
            case InteractiveInfoAtom.LINK_CustomShow:
            case InteractiveInfoAtom.LINK_OtherPresentation:
            case InteractiveInfoAtom.LINK_OtherFile:
                return HyperlinkType.FILE;
            default:
            case InteractiveInfoAtom.LINK_NULL:
                return HyperlinkType.NONE;
        }
    }

    //@Override
    public void LinkToEmail(String emailAddress) {
        InteractiveInfoAtom iia = info.GetInteractiveInfoAtom();
        iia.SetAction(InteractiveInfoAtom.ACTION_HYPERLINK);
        iia.SetJump(InteractiveInfoAtom.JUMP_NONE);
        iia.SetHyperlinkType(InteractiveInfoAtom.LINK_Url);
        exHyper.SetLinkURL("mailto:"+emailAddress);
        exHyper.SetLinkTitle(emailAddress);
        exHyper.SetLinkOptions(0x10);
    }

    //@Override
    public void LinkToUrl(String url) {
        InteractiveInfoAtom iia = info.GetInteractiveInfoAtom();
        iia.SetAction(InteractiveInfoAtom.ACTION_HYPERLINK);
        iia.SetJump(InteractiveInfoAtom.JUMP_NONE);
        iia.SetHyperlinkType(InteractiveInfoAtom.LINK_Url);
        exHyper.SetLinkURL(url);
        exHyper.SetLinkTitle(url);
        exHyper.SetLinkOptions(0x10);
    }

    //@Override
    public void LinkToSlide(Slide<HSLFShape,HSLFTextParagraph> slide) {
        //assert(slide instanceof HSLFSlide);
        HSLFSlide sl = (HSLFSlide)slide;
        int slideNum = slide.GetSlideNumber();
        String alias = "Slide "+slideNum;

        InteractiveInfoAtom iia = info.GetInteractiveInfoAtom();
        iia.SetAction(InteractiveInfoAtom.ACTION_HYPERLINK);
        iia.SetJump(InteractiveInfoAtom.JUMP_NONE);
        iia.SetHyperlinkType(InteractiveInfoAtom.LINK_SlideNumber);

        LinkToDocument(sl._getSheetNumber(),slideNum,alias,0x30);
    }

    //@Override
    public void LinkToNextSlide() {
        InteractiveInfoAtom iia = info.GetInteractiveInfoAtom();
        iia.SetAction(InteractiveInfoAtom.ACTION_JUMP);
        iia.SetJump(InteractiveInfoAtom.JUMP_NEXTSLIDE);
        iia.SetHyperlinkType(InteractiveInfoAtom.LINK_NextSlide);

        LinkToDocument(1,-1,"NEXT",0x10);
    }

    //@Override
    public void LinkToPreviousSlide() {
        InteractiveInfoAtom iia = info.GetInteractiveInfoAtom();
        iia.SetAction(InteractiveInfoAtom.ACTION_JUMP);
        iia.SetJump(InteractiveInfoAtom.JUMP_PREVIOUSSLIDE);
        iia.SetHyperlinkType(InteractiveInfoAtom.LINK_PreviousSlide);

        LinkToDocument(1,-1,"PREV",0x10);
    }

    //@Override
    public void LinkToFirstSlide() {
        InteractiveInfoAtom iia = info.GetInteractiveInfoAtom();
        iia.SetAction(InteractiveInfoAtom.ACTION_JUMP);
        iia.SetJump(InteractiveInfoAtom.JUMP_FIRSTSLIDE);
        iia.SetHyperlinkType(InteractiveInfoAtom.LINK_FirstSlide);

        LinkToDocument(1,-1,"FIRST",0x10);
    }

    //@Override
    public void LinkToLastSlide() {
        InteractiveInfoAtom iia = info.GetInteractiveInfoAtom();
        iia.SetAction(InteractiveInfoAtom.ACTION_JUMP);
        iia.SetJump(InteractiveInfoAtom.JUMP_LASTSLIDE);
        iia.SetHyperlinkType(InteractiveInfoAtom.LINK_LastSlide);

        LinkToDocument(1,-1,"LAST",0x10);
    }

    private void LinkToDocument(int sheetNumber, int slideNumber, String alias, int options) {
        exHyper.SetLinkURL(sheetNumber+","+slideNumber+","+alias);
        exHyper.SetLinkTitle(alias);
        exHyper.SetLinkOptions(options);
    }

    //@Override
    public String GetAddress() {
        return exHyper.GetLinkURL();
    }

    //@Override
    public void SetAddress(String str) {
        exHyper.SetLinkURL(str);
    }

    public int GetId() {
        return exHyper.GetExHyperlinkAtom().GetNumber();
    }

    //@Override
    public String GetLabel() {
        return exHyper.GetLinkTitle();
    }

    //@Override
    public void SetLabel(String label) {
        exHyper.SetLinkTitle(label);
    }

    /**
     * Gets the beginning character position
     *
     * @return the beginning character position
     */
    public int GetStartIndex() {
        return (txinfo == null) ? -1 : txinfo.GetStartIndex();
    }

    /**
     * Sets the beginning character position
     *
     * @param startIndex the beginning character position
     */
    public void SetStartIndex(int startIndex) {
        if (txinfo != null) {
            txinfo.SetStartIndex(startIndex);
        }
    }

    /**
     * Gets the ending character position
     *
     * @return the ending character position
     */
    public int GetEndIndex() {
        return (txinfo == null) ? -1 : txinfo.GetEndIndex();
    }

    /**
     * Sets the ending character position
     *
     * @param endIndex the ending character position
     */
    public void SetEndIndex(int endIndex) {
        if (txinfo != null) {
            txinfo.SetEndIndex(endIndex);
        }
    }

    /**
     * Find hyperlinks in a text shape
     *
     * @param shape  {@code TextRun} to lookup hyperlinks in
     * @return found hyperlinks or {@code null} if not found
     */
    public static List<HSLFHyperlink> Find(HSLFTextShape shape){
        return Find(shape.GetTextParagraphs());
    }

    /**
     * Find hyperlinks in a text paragraph
     *
     * @param paragraphs  List of {@code TextParagraph} to lookup hyperlinks
     * @return found hyperlinks
     */
    //@SuppressWarnings("resource")
    protected static List<HSLFHyperlink> Find(List<HSLFTextParagraph> paragraphs){
            List<HSLFHyperlink> lst = new List<HSLFHyperlink>();
        if (paragraphs == null || paragraphs.Count == 0) return lst;

        HSLFTextParagraph firstPara = paragraphs.Get(0);

        HSLFSlideShow ppt = firstPara.GetSheet().GetSlideShow();
        //document-level container which stores info about all links in a presentation
        ExObjList exobj = ppt.GetDocumentRecord().GetExObjList(false);
        if (exobj != null) {
            Record.Record[] records = firstPara.GetRecords();
            Find(records.ToList(), exobj, lst);
        }

        return lst;
    }

    /**
     * Find hyperlink assigned to the supplied shape
     *
     * @param shape  {@code Shape} to lookup hyperlink in
     * @return found hyperlink or {@code null}
     */
    //@SuppressWarnings("resource")
    protected static HSLFHyperlink Find(HSLFShape shape){
        HSLFSlideShow ppt = shape.GetSheet().GetSlideShow();
        //document-level container which stores info about all links in a presentation
        ExObjList exobj = ppt.GetDocumentRecord().GetExObjList(false);
        HSLFEscherClientDataRecord cldata = shape.GetClientData(false);

        if (exobj != null && cldata != null) {
                List<HSLFHyperlink> lst = new List<HSLFHyperlink>();
            Find(cldata.GetHSLFChildRecords(), exobj, lst);
            return lst.Count==0 ? null : lst.ElementAt(0);
        }

        return null;
    }

    private static void Find(List<Record.Record> records, ExObjList exobj, List<HSLFHyperlink> _out){
        List<Record.Record> iter = records.ListIterator();
        while (iter.hasNext()) {
            org.apache.poi.hslf.record.Record r = iter.next();
            // see if we have InteractiveInfo in the textrun's records
            if (!(r is InteractiveInfo)) {
                continue;
            }

            InteractiveInfo hldr = (InteractiveInfo)r;
            InteractiveInfoAtom info = hldr.GetInteractiveInfoAtom();
            if (info == null) {
                continue;
            }
            int id = info.GetHyperlinkID();
            ExHyperlink exHyper = exobj.Get(id);
            if (exHyper == null) {
                continue;
            }

            HSLFHyperlink link = new HSLFHyperlink(exHyper, hldr);
            _out.Add(link);

            if (iter.hasNext()) {
                r = iter.next();
                if (!(r is TxInteractiveInfoAtom)) {
                    iter.previous();
                    continue;
                }
                link.SetTextRunInfo((TxInteractiveInfoAtom)r);
            }
        }
    }
	}
}
