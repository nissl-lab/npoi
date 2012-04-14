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
namespace NPOI.XWPF.Model
{
    using System;
    using NPOI.XWPF.UserModel;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.Collections.Generic;
    using System.IO;

    /**
     * A .docx file can have no headers/footers, the same header/footer
     *  on each page, odd/even page footers, and optionally also 
     *  a different header/footer on the first page.
     * This class handles sorting out what there is, and giving you
     *  the right headers and footers for the document.
     */
    public class XWPFHeaderFooterPolicy
    {
        public static ST_HdrFtr DEFAULT = ST_HdrFtr.DEFAULT;
        public static ST_HdrFtr EVEN = ST_HdrFtr.EVEN;
        public static ST_HdrFtr FIRST = ST_HdrFtr.FIRST;

        private XWPFDocument doc;

        private XWPFHeader firstPageHeader;
        private XWPFFooter firstPageFooter;

        private XWPFHeader evenPageHeader;
        private XWPFFooter evenPageFooter;

        private XWPFHeader defaultHeader;
        private XWPFFooter defaultFooter;

        /**
         * Figures out the policy for the given document,
         *  and Creates any header and footer objects
         *  as required.
         */
        public XWPFHeaderFooterPolicy(XWPFDocument doc)
        {
            //this(doc, doc.GetDocument().GetBody().SectPr);
        }

        /**
         * Figures out the policy for the given document,
         *  and Creates any header and footer objects
         *  as required.
         */
        public XWPFHeaderFooterPolicy(XWPFDocument doc, CT_SectPr sectPr)
        {
            // Grab what headers and footers have been defined
            // For now, we don't care about different ranges, as it
            //  doesn't seem that .docx properly supports that
            //  feature of the file format yet
            this.doc = doc;
            //for(int i=0; i<sectPr.SizeOfHeaderReferenceArray(); i++) {
            //    // Get the header
            //    CTHdrFtrRef ref1 = sectPr.GetHeaderReferenceArray(i);
            //    POIXMLDocumentPart relatedPart = doc.GetRelationById(ref1.Id);
            //    XWPFHeader hdr = null;
            //    if (relatedPart != null && relatedPart is XWPFHeader) {
            //        hdr = (XWPFHeader) relatedPart;
            //    }
            //    // Assign it
            //    Enum type = ref1.Type;
            //    assignHeader(hdr, type);
            //}
            //for(int i=0; i<sectPr.SizeOfFooterReferenceArray(); i++) {
            //    // Get the footer
            //    CTHdrFtrRef ref1 = sectPr.GetFooterReferenceArray(i);
            //    POIXMLDocumentPart relatedPart = doc.GetRelationById(ref1.Id);
            //    XWPFFooter ftr = null;
            //    if (relatedPart != null && relatedPart is XWPFFooter)
            //    {
            //        ftr = (XWPFFooter) relatedPart;
            //    }
            //    // Assign it
            //    Enum type = ref1.Type;
            //    assignFooter(ftr, type);
            //}
            throw new NotImplementedException();
        }

        private void assignFooter(XWPFFooter ftr, ST_HdrFtr type)
        {
            if (type == ST_HdrFtr.FIRST)
            {
                firstPageFooter = ftr;
            }
            else if (type == ST_HdrFtr.EVEN)
            {
                evenPageFooter = ftr;
            }
            else
            {
                defaultFooter = ftr;
            }
        }

        private void assignHeader(XWPFHeader hdr, ST_HdrFtr type)
        {
            if (type == ST_HdrFtr.FIRST)
            {
                firstPageHeader = hdr;
            }
            else if (type == ST_HdrFtr.EVEN)
            {
                evenPageHeader = hdr;
            }
            else
            {
                defaultHeader = hdr;
            }
        }

        public XWPFHeader CreateHeader(Enum type)
        {
            return CreateHeader(type, null);
        }

        public XWPFHeader CreateHeader(Enum type, XWPFParagraph[] pars)
        {
            XWPFRelation relation = XWPFRelation.HEADER;
            String pStyle = "Header";
            int i = GetRelationIndex(relation);
            //HdrDocument hdrDoc = HdrDocument.Factory.NewInstance();
            //XWPFHeader wrapper = (XWPFHeader)doc.CreateRelationship(relation, XWPFFactory.Instance, i);

            //CTHdrFtr hdr = buildHdr(type, pStyle, wrapper, pars);
            //wrapper.HeaderFooter=(hdr);

            //OutputStream outputStream = wrapper.PackagePart.OutputStream;
            //hdrDoc.Hdr=(hdr);

            //XmlOptions xmlOptions = Commit(wrapper);

            //assignHeader(wrapper, type);
            //hdrDoc.Save(outputStream, xmlOptions);
            //outputStream.Close();
            //return wrapper;
            throw new NotImplementedException();
        }

        public XWPFFooter CreateFooter(Enum type)
        {
            return CreateFooter(type, null);
        }

        public XWPFFooter CreateFooter(Enum type, XWPFParagraph[] pars)
        {
            XWPFRelation relation = XWPFRelation.FOOTER;
            String pStyle = "Footer";
            int i = GetRelationIndex(relation);
            //FtrDocument ftrDoc = FtrDocument.Factory.NewInstance();
            //XWPFFooter wrapper = (XWPFFooter)doc.CreateRelationship(relation, XWPFFactory.Instance, i);

            //CTHdrFtr ftr = buildFtr(type, pStyle, wrapper, pars);
            //wrapper.HeaderFooter=(ftr);

            //OutputStream outputStream = wrapper.PackagePart.OutputStream;
            //ftrDoc.Ftr=(ftr);

            //XmlOptions xmlOptions = Commit(wrapper);

            //assignFooter(wrapper, type);
            //ftrDoc.Save(outputStream, xmlOptions);
            //outputStream.Close();
            //return wrapper;
            throw new NotImplementedException();
        }

        private int GetRelationIndex(XWPFRelation relation)
        {
            //List<POIXMLDocumentPart> relations = doc.Relations;
            //int i = 1;
            //for (Iterator<POIXMLDocumentPart> it = relations.Iterator(); it.HasNext() ; ) {
            //    POIXMLDocumentPart item = it.Next();
            //    if (item.PackageRelationship.RelationshipType.Equals(relation.Relation)) {
            //        i++;	
            //    }
            //}
            //return i;
            throw new NotImplementedException();
        }

        private CT_HdrFtr buildFtr(Enum type, String pStyle, XWPFHeaderFooter wrapper, XWPFParagraph[] pars)
        {
            //CTHdrFtr ftr = buildHdrFtr(pStyle, pars);				// MB 24 May 2010
            CT_HdrFtr ftr = buildHdrFtr(pStyle, pars, wrapper);		// MB 24 May 2010
            SetFooterReference(type, wrapper);
            return ftr;
        }

        private CT_HdrFtr buildHdr(Enum type, String pStyle, XWPFHeaderFooter wrapper, XWPFParagraph[] pars)
        {
            //CTHdrFtr hdr = buildHdrFtr(pStyle, pars);				// MB 24 May 2010
            CT_HdrFtr hdr = buildHdrFtr(pStyle, pars, wrapper);		// MB 24 May 2010
            SetHeaderReference(type, wrapper);
            return hdr;
        }

        private CT_HdrFtr buildHdrFtr(String pStyle, XWPFParagraph[] paragraphs)
        {
            //CTHdrFtr ftr = CTHdrFtr.Factory.NewInstance();
            //if (paragraphs != null) {
            //    for (int i = 0 ; i < paragraphs.Length ; i++) {
            //        CTP p = ftr.AddNewP();
            //        //ftr.PArray=(0, paragraphs[i].CTP);		// MB 23 May 2010
            //        ftr.PArray=(i, paragraphs[i].CTP);   	// MB 23 May 2010
            //    }
            //}
            //else {
            //    CTP p = ftr.AddNewP();
            //    byte[] rsidr = doc.Document.Body.GetPArray(0).RsidR;
            //    byte[] rsidrdefault = doc.Document.Body.GetPArray(0).RsidRDefault;
            //    p.RsidP=(rsidr);
            //    p.RsidRDefault=(rsidrdefault);
            //    CTPPr pPr = p.AddNewPPr();
            //    pPr.AddNewPStyle().Val=(pStyle);
            //}
            //return ftr;
            throw new NotImplementedException();
        }

        /**
         * MB 24 May 2010. Created this overloaded buildHdrFtr() method because testing demonstrated
         * that the XWPFFooter or XWPFHeader object returned by calls to the CreateHeader(int, XWPFParagraph[])
         * and CreateFooter(int, XWPFParagraph[]) methods or the GetXXXXXHeader/Footer methods where
         * headers or footers had been Added to a document since it had been Created/opened, returned
         * an object that Contained no XWPFParagraph objects even if the header/footer itself did contain
         * text. The reason was that this line of code; CTHdrFtr ftr = CTHdrFtr.Factory.NewInstance(); 
         * Created a brand new instance of the CTHDRFtr class which was then populated with data when
         * it should have recovered the CTHdrFtr object encapsulated within the XWPFHeaderFooter object
         * that had previoulsy been instantiated in the CreateHeader(int, XWPFParagraph[]) or 
         * CreateFooter(int, XWPFParagraph[]) methods.
         */
        private CT_HdrFtr buildHdrFtr(String pStyle, XWPFParagraph[] paragraphs, XWPFHeaderFooter wrapper)
        {
            CT_HdrFtr ftr = wrapper._getHdrFtr();
            //if (paragraphs != null) {
            //    for (int i = 0 ; i < paragraphs.Length ; i++) {
            //        CT_P p = ftr.AddNewP();
            //        ftr.PArray=(i, paragraphs[i].CTP);
            //    }
            //}
            //else {
            //    CTP p = ftr.AddNewP();
            //    byte[] rsidr = doc.Document.Body.GetPArray(0).RsidR;
            //    byte[] rsidrdefault = doc.Document.Body.GetPArray(0).RsidRDefault;
            //    p.RsidP=(rsidr);
            //    p.RsidRDefault=(rsidrdefault);
            //    CTPPr pPr = p.AddNewPPr();
            //    pPr.AddNewPStyle().Val=(pStyle);
            //}
            //return ftr;
            throw new NotImplementedException();
        }


        private void SetFooterReference(Enum type, XWPFHeaderFooter wrapper)
        {
            //CTHdrFtrRef ref1 = doc.Document.Body.SectPr.AddNewFooterReference();
            //ref1.Type=(type);
            //ref1.Id=(wrapper.PackageRelationship.Id);
            throw new NotImplementedException();
        }


        private void SetHeaderReference(Enum type, XWPFHeaderFooter wrapper)
        {
            //CTHdrFtrRef ref1 = doc.Document.Body.SectPr.AddNewHeaderReference();
            //ref1.Type=(type);
            //ref1.Id=(wrapper.PackageRelationship.Id);
            throw new NotImplementedException();
        }


        //private XmlOptions Commit(XWPFHeaderFooter wrapper) {
        //    XmlOptions xmlOptions = new XmlOptions(wrapper.DEFAULT_XML_OPTIONS);
        //    Dictionary<String, String> map = new Dictionary<String, String>();
        //    map.Put("http://schemas.Openxmlformats.org/officeDocument/2006/math", "m");
        //    map.Put("urn:schemas-microsoft-com:office:office", "o");
        //    map.Put("http://schemas.Openxmlformats.org/officeDocument/2006/relationships", "r");
        //    map.Put("urn:schemas-microsoft-com:vml", "v");
        //    map.Put("http://schemas.Openxmlformats.org/markup-compatibility/2006", "ve");
        //    map.Put("http://schemas.Openxmlformats.org/wordProcessingml/2006/main", "w");
        //    map.Put("urn:schemas-microsoft-com:office:word", "w10");
        //    map.Put("http://schemas.microsoft.com/office/word/2006/wordml", "wne");
        //    map.Put("http://schemas.Openxmlformats.org/drawingml/2006/wordProcessingDrawing", "wp");
        //    xmlOptions.SaveSuggestedPrefixes=(map);
        //    return xmlOptions;
        //}

        public XWPFHeader GetFirstPageHeader()
        {
            return firstPageHeader;
        }
        public XWPFFooter GetFirstPageFooter()
        {
            return firstPageFooter;
        }
        /**
         * Returns the odd page header. This is
         *  also the same as the default one...
         */
        public XWPFHeader GetOddPageHeader()
        {
            return defaultHeader;
        }
        /**
         * Returns the odd page footer. This is
         *  also the same as the default one...
         */
        public XWPFFooter GetOddPageFooter()
        {
            return defaultFooter;
        }
        public XWPFHeader GetEvenPageHeader()
        {
            return evenPageHeader;
        }
        public XWPFFooter GetEvenPageFooter()
        {
            return evenPageFooter;
        }
        public XWPFHeader GetDefaultHeader()
        {
            return defaultHeader;
        }
        public XWPFFooter GetDefaultFooter()
        {
            return defaultFooter;
        }

        /**
         * Get the header that applies to the given
         *  (1 based) page.
         * @param pageNumber The one based page number
         */
        public XWPFHeader GetHeader(int pageNumber)
        {
            if (pageNumber == 1 && firstPageHeader != null)
            {
                return firstPageHeader;
            }
            if (pageNumber % 2 == 0 && evenPageHeader != null)
            {
                return evenPageHeader;
            }
            return defaultHeader;
        }
        /**
         * Get the footer that applies to the given
         *  (1 based) page.
         * @param pageNumber The one based page number
         */
        public XWPFFooter GetFooter(int pageNumber)
        {
            if (pageNumber == 1 && firstPageFooter != null)
            {
                return firstPageFooter;
            }
            if (pageNumber % 2 == 0 && evenPageFooter != null)
            {
                return evenPageFooter;
            }
            return defaultFooter;
        }

        public void CreateWatermark(String text)
        {
            XWPFParagraph[] pars = new XWPFParagraph[1];
            try
            {
                pars[0] = GetWatermarkParagraph(text, 1);
                CreateHeader(DEFAULT, pars);
                pars[0] = GetWatermarkParagraph(text, 2);
                CreateHeader(FIRST, pars);
                pars[0] = GetWatermarkParagraph(text, 3);
                CreateHeader(EVEN, pars);
            }
            catch (IOException e)
            {
                // TODO Auto-generated catch block
                //e.PrintStackTrace();
                Console.Write(e.StackTrace);
            }
        }

        /*
         * This is the default Watermark paragraph; the only variable is the text message
         * TODO: manage all the other variables
         */
        private XWPFParagraph GetWatermarkParagraph(String text, int idx)
        {
            //CTP p = CTP.Factory.NewInstance();
            //byte[] rsidr = doc.Document.Body.GetPArray(0).RsidR;
            //byte[] rsidrdefault = doc.Document.Body.GetPArray(0).RsidRDefault;
            //p.RsidP=(rsidr);
            //p.RsidRDefault=(rsidrdefault);
            //CTPPr pPr = p.AddNewPPr();
            //pPr.AddNewPStyle().Val=("Header");
            //// start watermark paragraph
            //CTR r = p.AddNewR();
            //CTRPr rPr = r.AddNewRPr();
            //rPr.AddNewNoProof();
            //CTPicture pict = r.AddNewPict();
            //CTGroup group = CTGroup.Factory.NewInstance();
            //CTShapetype shapetype = group.AddNewShapetype();
            //shapetype.Id=("_x0000_t136");
            //shapetype.Coordsize=("1600,21600");
            //shapetype.Spt=(136);
            //shapetype.Adj=("10800");
            //shapetype.Path2=("m@7,0l@8,0m@5,21600l@6,21600e");
            //CTFormulas formulas = shapetype.AddNewFormulas();
            //formulas.AddNewF().Eqn=("sum #0 0 10800");
            //formulas.AddNewF().Eqn=("prod #0 2 1");
            //formulas.AddNewF().Eqn=("sum 21600 0 @1");
            //formulas.AddNewF().Eqn=("sum 0 0 @2");
            //formulas.AddNewF().Eqn=("sum 21600 0 @3");
            //formulas.AddNewF().Eqn=("if @0 @3 0");
            //formulas.AddNewF().Eqn=("if @0 21600 @1");
            //formulas.AddNewF().Eqn=("if @0 0 @2");
            //formulas.AddNewF().Eqn=("if @0 @4 21600");
            //formulas.AddNewF().Eqn=("mid @5 @6");
            //formulas.AddNewF().Eqn=("mid @8 @5");
            //formulas.AddNewF().Eqn=("mid @7 @8");
            //formulas.AddNewF().Eqn=("mid @6 @7");
            //formulas.AddNewF().Eqn=("sum @6 0 @5");
            //CTPath path = shapetype.AddNewPath();
            //path.Textpathok=(STTrueFalse.T);
            //path.Connecttype=(STConnectType.CUSTOM);
            //path.Connectlocs=("@9,0;@10,10800;@11,21600;@12,10800");
            //path.Connectangles=("270,180,90,0");
            //CTTextPath shapeTypeTextPath = shapetype.AddNewTextpath();
            //shapeTypeTextPath.On=(STTrueFalse.T);
            //shapeTypeTextPath.Fitshape=(STTrueFalse.T);
            //CTHandles handles = shapetype.AddNewHandles();
            //CTH h = handles.AddNewH();
            //h.Position=("#0,bottomRight");
            //h.Xrange=("6629,14971");
            //CTLock lock = shapetype.AddNewLock();
            //lock.Ext=(STExt.EDIT);
            //CTShape shape = group.AddNewShape();
            //shape.Id=("PowerPlusWaterMarkObject" + idx);
            //shape.Spid=("_x0000_s102" + (4+idx));
            //shape.Type=("#_x0000_t136");
            //shape.Style=("position:absolute;margin-left:0;margin-top:0;width:415pt;height:207.5pt;z-index:-251654144;mso-wrap-edited:f;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin");
            //shape.Wrapcoords=("616 5068 390 16297 39 16921 -39 17155 7265 17545 7186 17467 -39 17467 18904 17467 10507 17467 8710 17545 18904 17077 18787 16843 18358 16297 18279 12554 19178 12476 20701 11774 20779 11228 21131 10059 21248 8811 21248 7563 20975 6316 20935 5380 19490 5146 14022 5068 2616 5068");
            //shape.Fillcolor=("black");
            //shape.Stroked=(STTrueFalse.FALSE);
            //CTTextPath shapeTextPath = shape.AddNewTextpath();
            //shapeTextPath.Style=("font-family:&quot;Cambria&quot;;font-size:1pt");
            //shapeTextPath.String=(text);
            //pict.Set(group);
            //// end watermark paragraph
            //return new XWPFParagraph(p, doc);
            throw new NotImplementedException();
        }
    }

}