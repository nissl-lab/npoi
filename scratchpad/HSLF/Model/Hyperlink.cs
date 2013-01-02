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

using System.Collections;
using NPOI.DDF;
namespace NPOI.HSLF.Model;

using NPOI.HSLF.record.*;
using NPOI.HSLF.usermodel.SlideShow;
using NPOI.ddf.EscherContainerRecord;
using NPOI.ddf.EscherRecord;
using NPOI.ddf.EscherClientDataRecord;





/**
 * Represents a hyperlink in a PowerPoint document
 *
 * @author Yegor Kozlov
 */
public class Hyperlink {
    public const byte LINK_NEXTSLIDE = InteractiveInfoAtom.LINK_NextSlide;
    public const byte LINK_PREVIOUSSLIDE = InteractiveInfoAtom.LINK_PreviousSlide;
    public const byte LINK_FIRSTSLIDE = InteractiveInfoAtom.LINK_FirstSlide;
    public const byte LINK_LASTSLIDE = InteractiveInfoAtom.LINK_LastSlide;
    public const byte LINK_URL = InteractiveInfoAtom.LINK_Url;
    public const byte LINK_NULL = InteractiveInfoAtom.LINK_NULL;

    private int id=-1;
    private int type;
    private String Address;
    private String title;
    private int startIndex, endIndex;

    /**
     * Gets the type of the hyperlink action.
     * Must be a <code>LINK_*</code>  constant</code>
     *
     * @return the hyperlink URL
     * @see InteractiveInfoAtom
     */
    public int GetType() {
        return type;
    }

    public void SetType(int val) {
        type = val;
        switch(type){
            case LINK_NEXTSLIDE:
                title = "NEXT";
                Address = "1,-1,NEXT";
                break;
            case LINK_PREVIOUSSLIDE:
                title = "PREV";
                Address = "1,-1,PREV";
                break;
            case LINK_FIRSTSLIDE:
                title = "FIRST";
                Address = "1,-1,FIRST";
                break;
            case LINK_LASTSLIDE:
                title = "LAST";
                Address = "1,-1,LAST";
                break;
            default:
                title = "";
                Address = "";
                break;
        }
    }

    /**
     * Gets the hyperlink URL
     *
     * @return the hyperlink URL
     */
    public String GetAddress() {
        return Address;
    }

    public void SetAddress(String str) {
        Address = str;
    }

    public int GetId() {
        return id;
    }

    public void SetId(int id) {
        this.id = id;
    }

    /**
     * Gets the hyperlink user-friendly title (if different from URL)
     *
     * @return the  hyperlink user-friendly title
     */
    public String GetTitle() {
        return title;
    }

    public void SetTitle(String str) {
        title = str;
    }

    /**
     * Gets the beginning character position
     *
     * @return the beginning character position
     */
    public int GetStartIndex() {
        return startIndex;
    }

    /**
     * Gets the ending character position
     *
     * @return the ending character position
     */
    public int GetEndIndex() {
        return endIndex;
    }

    /**
     * Find hyperlinks in a text run
     *
     * @param run  <code>TextRun</code> to lookup hyperlinks in
     * @return found hyperlinks or <code>null</code> if not found
     */
    protected static Hyperlink[] Find(TextRun Run){
        ArrayList lst = new ArrayList();
        SlideShow ppt = Run.Sheet.GetSlideShow();
        //document-level Container which stores info about all links in a presentation
        ExObjList exobj = ppt.GetDocumentRecord().GetExObjList();
        if (exobj == null) {
            return null;
        }
        Record[] records = Run._records;
        if(records != null) Find(records, exobj, lst);

        Hyperlink[] links = null;
        if (lst.Count > 0){
            links = new Hyperlink[lst.Count];
            lst.ToArray(links);
        }
        return links;
    }

    /**
     * Find hyperlink assigned to the supplied shape
     *
     * @param shape  <code>Shape</code> to lookup hyperlink in
     * @return found hyperlink or <code>null</code>
     */
    protected static Hyperlink Find(Shape shape){
        ArrayList lst = new ArrayList();
        SlideShow ppt = shape.Sheet.GetSlideShow();
        //document-level Container which stores info about all links in a presentation
        ExObjList exobj = ppt.GetDocumentRecord().GetExObjList();
        if (exobj == null) {
            return null;
        }

        EscherContainerRecord spContainer = shape.GetSpContainer();
        for (Iterator<EscherRecord> it = spContainer.GetChildIterator(); it.HasNext(); ) {
            EscherRecord obj = it.next();
            if (obj.GetRecordId() ==  EscherClientDataRecord.RECORD_ID){
                byte[] data = obj.Serialize();
                Record[] records = Record.FindChildRecords(data, 8, data.Length-8);
                if(records != null) Find(records, exobj, lst);
            }
        }

        return lst.Count == 1 ? (Hyperlink)lst.Get(0) : null;
    }

    private static void Find(Record[] records, ExObjList exobj, List out1){
        for (int i = 0; i < records.Length; i++) {
            //see if we have InteractiveInfo in the textRun's records
            if( records[i] is InteractiveInfo){
                InteractiveInfo hldr = (InteractiveInfo)records[i];
                InteractiveInfoAtom info = hldr.GetInteractiveInfoAtom();
                int id = info.GetHyperlinkID();
                ExHyperlink linkRecord = exobj.Get(id);
                if (linkRecord != null){
                    Hyperlink link = new Hyperlink();
                    link.title = linkRecord.GetLinkTitle();
                    link.Address = linkRecord.GetLinkURL();
                    link.type = info.GetAction();

                    if (++i < records.Length && records[i] is TxInteractiveInfoAtom){
                        TxInteractiveInfoAtom txinfo = (TxInteractiveInfoAtom)records[i];
                        link.startIndex = txinfo.GetStartIndex();
                        link.endIndex = txinfo.GetEndIndex();
                    }
                    out1.Add(link);
                }
            }
        }
    }
}





