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

using NPOI.HSLF.record.*;
using NPOI.HSLF.usermodel.SlideShow;

/**
 * Header / Footer Settings.
 *
 * You can Get these on slides, or across all notes
 *
 * @author Yegor Kozlov
 */
public class HeadersFooters {

    private HeadersFootersContainer _Container;
    private bool _newRecord;
    private SlideShow _ppt;
    private Sheet _sheet;
    private bool _ppt2007;


    public HeadersFooters(HeadersFootersContainer rec, SlideShow ppt, bool newRecord, bool IsPpt2007){
        _Container = rec;
        _newRecord = newRecord;
        _ppt = ppt;
        _ppt2007 = isPpt2007;
    }

    public HeadersFooters(HeadersFootersContainer rec, Sheet sheet, bool newRecord, bool IsPpt2007){
        _Container = rec;
        _newRecord = newRecord;
        _sheet = sheet;
        _ppt2007 = isPpt2007;
    }

    /**
     * Headers's text
     *
     * @return Headers's text
     */
    public String GetHeaderText(){
        CString cs = _Container == null ? null : _Container.GetHeaderAtom();
        return GetPlaceholderText(OEPlaceholderAtom.MasterHeader, cs);
    }

    /**
     * Sets headers's text
     *
     * @param text headers's text
     */
    public void SetHeaderText(String text){
        if(_newRecord) attach();

        SetHeaderVisible(true);
        CString cs = _Container.GetHeaderAtom();
        if(cs == null) cs = _Container.AddHeaderAtom();

        cs.SetText(text);
    }

    /**
     * Footer's text
     *
     * @return Footer's text
     */
    public String GetFooterText(){
        CString cs = _Container == null ? null : _Container.GetFooterAtom();
        return GetPlaceholderText(OEPlaceholderAtom.MasterFooter, cs);
    }

    /**
     * Sets footers's text
     *
     * @param text footers's text
     */
    public void SetFootersText(String text){
        if(_newRecord) attach();

        SetFooterVisible(true);
        CString cs = _Container.GetFooterAtom();
        if(cs == null) cs = _Container.AddFooterAtom();

        cs.SetText(text);
    }

    /**
     * This is the date that the user wants in the footers, instead of today's date.
     *
     * @return custom user date
     */
    public String GetDateTimeText(){
        CString cs = _Container == null ? null : _Container.GetUserDateAtom();
        return GetPlaceholderText(OEPlaceholderAtom.MasterDate, cs);
    }

    /**
     * Sets custom user date to be displayed instead of today's date.
     *
     * @param text custom user date
     */
    public void SetDateTimeText(String text){
        if(_newRecord) attach();

        SetUserDateVisible(true);
        SetDateTimeVisible(true);
        CString cs = _Container.GetUserDateAtom();
        if(cs == null) cs = _Container.AddUserDateAtom();

        cs.SetText(text);
    }

    /**
     * whether the footer text is displayed.
     */
    public bool IsFooterVisible(){
        return isVisible(HeadersFootersAtom.fHasFooter, OEPlaceholderAtom.MasterFooter);
    }

    /**
     * whether the footer text is displayed.
     */
    public void SetFooterVisible(bool flag){
        if(_newRecord) attach();
        _Container.GetHeadersFootersAtom().SetFlag(HeadersFootersAtom.fHasFooter, flag);
    }

    /**
     * whether the header text is displayed.
     */
    public bool IsHeaderVisible(){
        return isVisible(HeadersFootersAtom.fHasHeader, OEPlaceholderAtom.MasterHeader);
    }

    /**
     * whether the header text is displayed.
     */
    public void SetHeaderVisible(bool flag){
        if(_newRecord) attach();
        _Container.GetHeadersFootersAtom().SetFlag(HeadersFootersAtom.fHasHeader, flag);
    }

    /**
     * whether the date is displayed in the footer.
     */
    public bool IsDateTimeVisible(){
        return isVisible(HeadersFootersAtom.fHasDate, OEPlaceholderAtom.MasterDate);
    }

    /**
     * whether the date is displayed in the footer.
     */
    public void SetDateTimeVisible(bool flag){
        if(_newRecord) attach();
        _Container.GetHeadersFootersAtom().SetFlag(HeadersFootersAtom.fHasDate, flag);
    }

    /**
     * whether the custom user date is used instead of today's date.
     */
    public bool IsUserDateVisible(){
        return isVisible(HeadersFootersAtom.fHasUserDate, OEPlaceholderAtom.MasterDate);
    }

    /**
     * whether the date is displayed in the footer.
     */
    public void SetUserDateVisible(bool flag){
        if(_newRecord) attach();
        _Container.GetHeadersFootersAtom().SetFlag(HeadersFootersAtom.fHasUserDate, flag);
    }

    /**
     * whether the slide number is displayed in the footer.
     */
    public bool IsSlideNumberVisible(){
        return isVisible(HeadersFootersAtom.fHasSlideNumber, OEPlaceholderAtom.MasterSlideNumber);
    }

    /**
     * whether the slide number is displayed in the footer.
     */
    public void SetSlideNumberVisible(bool flag){
        if(_newRecord) attach();
        _Container.GetHeadersFootersAtom().SetFlag(HeadersFootersAtom.fHasSlideNumber, flag);
    }

    /**
     *  An integer that specifies the format ID to be used to style the datetime.
     *
     * @return an integer that specifies the format ID to be used to style the datetime.
     */
    public int GetDateTimeFormat(){
        return _Container.GetHeadersFootersAtom().GetFormatId();
    }

    /**
     *  An integer that specifies the format ID to be used to style the datetime.
     *
     * @param formatId an integer that specifies the format ID to be used to style the datetime.
     */
    public void SetDateTimeFormat(int formatId){
        if(_newRecord) attach();
        _Container.GetHeadersFootersAtom().SetFormatId(formatId);
    }

    /**
     * Attach this HeadersFootersContainer to the parent Document record
     */
    private void attach(){
        Document doc = _ppt.GetDocumentRecord();
        Record[] ch = doc.GetChildRecords();
        Record lst = null;
        for (int i=0; i < ch.Length; i++){
            if(ch[i].GetRecordType() == RecordTypes.List.typeID){
                lst = ch[i];
                break;
            }
        }
        doc.AddChildAfter(_Container, lst);
        _newRecord = false;
    }

    private bool IsVisible(int flag, int placeholderId){
        bool visible;
        if(_ppt2007){
            Sheet master = _sheet != null ? _sheet : _ppt.GetSlidesMasters()[0];
            TextShape placeholder = master.GetPlaceholder(placeholderId);
            visible = placeholder != null && placeholder.GetText() != null;
        } else {
            visible = _Container.GetHeadersFootersAtom().GetFlag(flag);
        }
        return visible;
    }

    private String GetPlaceholderText(int placeholderId, CString cs){
        String text = null;
        if(_ppt2007){
            Sheet master = _sheet != null ? _sheet : _ppt.GetSlidesMasters()[0];
            TextShape placeholder = master.GetPlaceholder(placeholderId);
            if(placeholder != null) text = placeholder.GetText();

            //default text in master placeholders is not visible
            if("*".Equals(text)) text = null;
        } else {
            text = cs == null ? null : cs.GetText();
        }
        return text;
    }

}





