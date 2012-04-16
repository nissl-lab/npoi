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
namespace NPOI.XWPF.UserModel
{
    using System;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.Collections.Generic;
    using System.Text;
    using System.Xml;
    using System.IO;
    /**
     * XWPFrun.object defines a region of text with a common Set of properties
     *
     * @author Yegor Kozlov
     */
    public class XWPFRun
    {
        private CT_R run;
        private String pictureText;
        private XWPFParagraph paragraph;
        private List<XWPFPicture> pictures;

        /**
         * @param r the CT_R bean which holds the run.attributes
         * @param p the parent paragraph
         */
        public XWPFRun(CT_R r, XWPFParagraph p)
        {
            this.run = r;
            this.paragraph = p;

            /**
             * reserve already occupied Drawing ids, so reserving new ids later will
             * not corrupt the document
             */
            //List<CT_Drawing> DrawingList = r.DrawingList;
            //foreach (CT_Drawing ctDrawing in DrawingList) {
            //    List<CT_Anchor> anchorList = ctDrawing.AnchorList;
            //    foreach (CT_Anchor anchor in anchorList) {
            //        if (anchor.DocPr != null) {
            //            GetDocument().DrawingIdManager.Reserve(anchor.DocPr.Id);
            //        }
            //    }
            //    List<CT_Inline> inlineList = ctDrawing.InlineList;
            //    foreach (CT_Inline inline in inlineList) {
            //        if (inline.DocPr != null) {
            //            GetDocument().DrawingIdManager.Reserve(inline.DocPr.Id);
            //        }
            //    }
            //}

            //// Look for any text in any of our pictures or Drawings
            //StringBuilder text = new StringBuilder();
            //List<XmlObject> pictTextObjs = new List<XmlObject>();
            //pictTextObjs.AddAll(r.PictList);
            //pictTextObjs.AddAll(drawingList);
            //foreach(XmlObject o in pictTextObjs) {
            //    XmlObject[] t = o.SelectPath("declare namespace w='http://schemas.Openxmlformats.org/wordProcessingml/2006/main' .//w:t");
            //    for (int m = 0; m < t.Length; m++) {
            //        NodeList kids = t[m].DomNode.ChildNodes;
            //        for (int n = 0; n < kids.Length; n++) {
            //            if (kids.Item(n) is Text) {
            //                if(text.Length() > 0)
            //                    text.Append("\n");
            //                text.Append(kids.Item(n).NodeValue);
            //            }
            //        }
            //    }
            //}
            //pictureText = text.ToString();

            //// Do we have any embedded pictures?
            //// (They're a different CT_Picture, under the Drawingml namespace)
            //pictures = new List<XWPFPicture>();
            //foreach(XmlObject o in pictTextObjs) {
            //    foreach(CT_Picture pict in GetCT_Pictures(o)) {
            //        XWPFPicture picture = new XWPFPicture(pict, this);
            //        pictures.Add(picture);
            //    }
            //}
        }

        private List<CT_Picture> GetCT_Pictures(/*XmlObject*/XmlElement o)
        {
            /*List<CT_Picture> pictures = new List<CT_Picture>(); 
            XmlObject[] picts = o.SelectPath("declare namespace pic='"+CT_Picture.type.Name.NamespaceURI+"' .//pic:pic");
            foreach(XmlObject pict in picts) {
                if(pict is XmlAnyTypeImpl) {
                    // Pesky XmlBeans bug - see Bugzilla #49934
                    try {
                        pict = CT_Picture.Factory.Parse( pict.ToString() );
                    } catch(XmlException e) {
                        throw new POIXMLException(e);
                    }
                }
                if(pict is CT_Picture) {
                    pictures.Add((CT_Picture)pict);
                }
            }
            return pictures;*/
            throw new NotImplementedException();
        }

        /**
         * Get the currently used CT_R object
         * @return CT_R object
         */

        public CT_R GetCT_R()
        {
            //return run.
            throw new NotImplementedException();
        }

        /**
         * Get the currenty referenced paragraph object
         * @return current paragraph
         */
        public XWPFParagraph GetParagraph()
        {
            return paragraph;
        }

        /**
         * @return The {@link XWPFDocument} instance, this run.belongs to, or
         *         <code>null</code> if parent structure (paragraph > document) is not properly Set.
         */
        public XWPFDocument GetDocument()
        {
            if (paragraph != null)
            {
                return paragraph.GetDocument();
            }
            return null;
        }

        /**
         * For isBold, isItalic etc
         */
        private bool IsCTOnOff(CT_OnOff onoff)
        {
            //if(! onoff.IsSetVal())
            //    return true;
            //if(onoff.Val == ST_OnOff.on)
            //    return true;
            //if(onoff.Val == ST_OnOff.@true)
            //    return true;
            //return false;
            throw new NotImplementedException();
        }

        /**
         * Whether the bold property shall be applied to all non-complex script
         * characters in the contents of this run.when displayed in a document
         *
         * @return <code>true</code> if the bold property is applied
         */
        public bool IsBold()
        {
            //CT_RPr pr = run.RPr;
            //if(pr == null || !pr.IsSetB()) {
            //    return false;
            //}
            //return isCTOnOff(pr.B);
            throw new NotImplementedException();
        }

        /**
         * Whether the bold property shall be applied to all non-complex script
         * characters in the contents of this run.when displayed in a document. 
         * <p>
         * This formatting property is a toggle property, which specifies that its
         * behavior differs between its use within a style defInition and its use as
         * direct formatting. When used as part of a style defInition, Setting this
         * property shall toggle the current state of that property as specified up
         * to this point in the hierarchy (i.e. applied to not applied, and vice
         * versa). Setting it to <code>false</code> (or an equivalent) shall
         * result in the current Setting remaining unChanged. However, when used as
         * direct formatting, Setting this property to true or false shall Set the
         * absolute state of the resulting property.
         * </p>
         * <p>
         * If this element is not present, the default value is to leave the
         * formatting applied at previous level in the style hierarchy. If this
         * element is never applied in the style hierarchy, then bold shall not be
         * applied to non-complex script characters.
         * </p>
         *
         * @param value <code>true</code> if the bold property is applied to
         *              this run
         */
        public void SetBold(bool value)
        {
            //CT_RPr pr = run.IsSetRPr() ? run.RPr : run.AddNewRPr();
            //CT_OnOff bold = pr.IsSetB() ? pr.B : pr.AddNewB();
            //bold.Val=(value ? STOnOff.TRUE : STOnOff.FALSE);
            throw new NotImplementedException();
        }

        /**
         * Return the string content of this text run
         *
         * @return the text of this text run.or <code>null</code> if not Set
         */
        public String GetText(int pos)
        {
            //return run.SizeOfTArray() == 0 ? null : run.GetTArray(pos)
            //        .StringValue;
            throw new NotImplementedException();
        }

        /**
         * Returns text embedded in pictures
         */
        public String GetPictureText()
        {
            return pictureText;
        }

        /**
         * Sets the text of this text run
         *
         * @param value the literal text which shall be displayed in the document
         */
        public void SetText(String value)
        {
            //SetText(value,run.TList.Size());
            SetText(value, run.Items.Length);
        }

        /**
         * Sets the text of this text run.in the 
         *
         * @param value the literal text which shall be displayed in the document
         * @param pos - position in the text array (NB: 0 based)
         */
        public void SetText(String value, int pos)
        {
            int length = run.Items.Length;
            if (pos > length) throw new IndexOutOfRangeException("Value too large for the parameter position in XWPFrun.Text=(String value,int pos)");
            CT_Text t = (pos < length && pos >= 0) ? run.Items[(pos)] as CT_Text : run.AddNewT();
            t.Value =(value);
            preserveSpaces(t);
        }

        /**
         * Whether the italic property should be applied to all non-complex script
         * characters in the contents of this run.when displayed in a document.
         *
         * @return <code>true</code> if the italic property is applied
         */
        public bool IsItalic()
        {
            //CT_RPr pr = run.RPr;
            //if(pr == null || !pr.IsSetI())
            //    return false;
            //return isCT_OnOff(pr.I);
            throw new NotImplementedException();
        }

        /**
         * Whether the bold property shall be applied to all non-complex script
         * characters in the contents of this run.when displayed in a document
         * <p/>
         * <p/>
         * This formatting property is a toggle property, which specifies that its
         * behavior differs between its use within a style defInition and its use as
         * direct formatting. When used as part of a style defInition, Setting this
         * property shall toggle the current state of that property as specified up
         * to this point in the hierarchy (i.e. applied to not applied, and vice
         * versa). Setting it to <code>false</code> (or an equivalent) shall
         * result in the current Setting remaining unChanged. However, when used as
         * direct formatting, Setting this property to true or false shall Set the
         * absolute state of the resulting property.
         * </p>
         * <p/>
         * If this element is not present, the default value is to leave the
         * formatting applied at previous level in the style hierarchy. If this
         * element is never applied in the style hierarchy, then bold shall not be
         * applied to non-complex script characters.
         * </p>
         *
         * @param value <code>true</code> if the italic property is applied to
         *              this run
         */
        public void SetItalic(bool value)
        {
            //CT_RPr pr = run.IsSetRPr() ? run.RPr : run.AddNewRPr();
            //CT_OnOff italic = pr.IsSetI() ? pr.I : pr.AddNewI();
            //italic.Val=(value ? STOnOff.TRUE : STOnOff.FALSE);
            throw new NotImplementedException();
        }

        /**
         * Specifies that the contents of this run.should be displayed along with an
         * underline appearing directly below the character heigh
         *
         * @return the Underline pattern Applyed to this run
         * @see UnderlinePatterns
         */
        public UnderlinePatterns GetUnderline()
        {
            //CT_RPr pr = run.RPr;
            //return (pr != null && pr.IsSetU()) ? UnderlinePatterns.ValueOf(pr
            //    .U.Val.IntValue()) : UnderlinePatterns.NONE;
            throw new NotImplementedException();
        }

        /**
         * Specifies that the contents of this run.should be displayed along with an
         * underline appearing directly below the character heigh
         * <p/>
         * <p/>
         * If this element is not present, the default value is to leave the
         * formatting applied at previous level in the style hierarchy. If this
         * element is never applied in the style hierarchy, then an underline shall
         * not be applied to the contents of this run.
         * </p>
         *
         * @param value -
         *              underline type
         * @see UnderlinePatterns : all possible patterns that could be applied
         */
        public void SetUnderline(UnderlinePatterns value)
        {
            //CT_RPr pr = run.IsSetRPr() ? run.RPr : run.AddNewRPr();
            //CTUnderline underline = (pr.U == null) ? pr.AddNewU() : pr.U;
            //underline.Val=(STUnderline.Enum.ForInt(value.Value));
            throw new NotImplementedException();
        }

        /**
         * Specifies that the contents of this run.shall be displayed with a single
         * horizontal line through the center of the line.
         *
         * @return <code>true</code> if the strike property is applied
         */
        public bool IsStrike()
        {
            //CT_RPr pr = run.RPr;
            //if(pr == null || !pr.IsSetStrike())
            //    return false;
            //return isCT_OnOff(pr.Strike);
            throw new NotImplementedException();
        }

        /**
         * Specifies that the contents of this run.shall be displayed with a single
         * horizontal line through the center of the line.
         * <p/>
         * This formatting property is a toggle property, which specifies that its
         * behavior differs between its use within a style defInition and its use as
         * direct formatting. When used as part of a style defInition, Setting this
         * property shall toggle the current state of that property as specified up
         * to this point in the hierarchy (i.e. applied to not applied, and vice
         * versa). Setting it to false (or an equivalent) shall result in the
         * current Setting remaining unChanged. However, when used as direct
         * formatting, Setting this property to true or false shall Set the absolute
         * state of the resulting property.
         * </p>
         * <p/>
         * If this element is not present, the default value is to leave the
         * formatting applied at previous level in the style hierarchy. If this
         * element is never applied in the style hierarchy, then strikethrough shall
         * not be applied to the contents of this run.
         * </p>
         *
         * @param value <code>true</code> if the strike property is applied to
         *              this run
         */
        public void SetStrike(bool value)
        {
            //CT_RPr pr = run.IsSetRPr() ? run.RPr : run.AddNewRPr();
            //CT_OnOff strike = pr.IsSetStrike() ? pr.Strike : pr.AddNewStrike();
            //strike.Val=(value ? STOnOff.TRUE : STOnOff.FALSE);
            throw new NotImplementedException();
        }

        /**
         * Specifies the alignment which shall be applied to the contents of this
         * run.in relation to the default appearance of the run.s text.
         * This allows the text to be repositioned as subscript or superscript without
         * altering the font size of the run.properties.
         *
         * @return VerticalAlign
         * @see VerticalAlign all possible value that could be Applyed to this run
         */
        public VerticalAlign GetSubscript()
        {
            //CT_RPr pr = run.RPr;
            //return (pr != null && pr.IsSetVertAlign()) ? VerticalAlign.ValueOf(pr.VertAlign.Val.IntValue()) : VerticalAlign.BASELINE;
            throw new NotImplementedException();
        }

        /**
         * Specifies the alignment which shall be applied to the contents of this
         * run.in relation to the default appearance of the run.s text. This allows
         * the text to be repositioned as subscript or superscript without altering
         * the font size of the run.properties.
         * <p/>
         * If this element is not present, the default value is to leave the
         * formatting applied at previous level in the style hierarchy. If this
         * element is never applied in the style hierarchy, then the text shall not
         * be subscript or superscript relative to the default baseline location for
         * the contents of this run.
         * </p>
         *
         * @param valign
         * @see VerticalAlign
         */
        public void SetSubscript(VerticalAlign valign)
        {
            //CT_RPr pr = run.IsSetRPr() ? run.RPr : run.AddNewRPr();
            //CTVerticalAlignrun.ctValign = pr.IsSetVertAlign() ? pr.VertAlign : pr.AddNewVertAlign();
            //ctValign.Val=(STVerticalAlignrun.Enum.ForInt(valign.Value));
            throw new NotImplementedException();
        }

        /**
         * Specifies the fonts which shall be used to display the text contents of
         * this run. Specifies a font which shall be used to format all characters
         * in the ASCII range (0 - 127) within the parent run
         *
         * @return a string representing the font family
         */
        public String GetFontFamily()
        {
            //CT_RPr pr = run.RPr;
            //return (pr != null && pr.IsSetRFonts()) ? pr.RFonts.Ascii
            //        : null;
            throw new NotImplementedException();
        }

        /**
         * Specifies the fonts which shall be used to display the text contents of
         * this run. Specifies a font which shall be used to format all characters
         * in the ASCII range (0 - 127) within the parent run
         *
         * @param fontFamily
         */
        public void SetFontFamily(String fontFamily)
        {
            //CT_RPr pr = run.RPr;
            //CTFonts fonts = pr.IsSetRFonts() ? pr.RFonts : pr.AddNewRFonts();
            //fonts.Ascii=(fontFamily);
            throw new NotImplementedException();
        }

        /**
         * Specifies the font size which shall be applied to all non complex script
         * characters in the contents of this run.when displayed.
         *
         * @return value representing the font size
         */
        public int GetFontSize()
        {
            //CT_RPr pr = run.RPr;
            //return (pr != null && pr.IsSetSz()) ? pr.Sz.Val.Divide(new Bigint("2")).intValue() : -1;
            throw new NotImplementedException();
        }

        /**
         * Specifies the font size which shall be applied to all non complex script
         * characters in the contents of this run.when displayed.
         * <p/>
         * If this element is not present, the default value is to leave the value
         * applied at previous level in the style hierarchy. If this element is
         * never applied in the style hierarchy, then any appropriate font size may
         * be used for non complex script characters.
         * </p>
         *
         * @param size
         */
        public void SetFontSize(int size)
        {
            //Bigint bint=new Bigint(""+size);
            //CT_RPr pr = run.IsSetRPr() ? run.RPr : run.AddNewRPr();
            //CTHpsMeasure ctSize = pr.IsSetSz() ? pr.Sz : pr.AddNewSz();
            //ctSize.Val=(bint.Multiply(new Bigint("2")));
            throw new NotImplementedException();
        }

        /**
         * This element specifies the amount by which text shall be raised or
         * lowered for this run.in relation to the default baseline of the
         * surrounding non-positioned text. This allows the text to be repositioned
         * without altering the font size of the contents.
         *
         * @return a big integer representing the amount of text shall be "moved"
         */
        public int GetTextPosition()
        {
            //CT_RPr pr = run.RPr;
            //return (pr != null && pr.IsSetPosition()) ? pr.Position.Val.IntValue()
            //        : -1;
            throw new NotImplementedException();
        }

        /**
         * This element specifies the amount by which text shall be raised or
         * lowered for this run.in relation to the default baseline of the
         * surrounding non-positioned text. This allows the text to be repositioned
         * without altering the font size of the contents.
         * <p/>
         * If the val attribute is positive, then the parent run.shall be raised
         * above the baseline of the surrounding text by the specified number of
         * half-points. If the val attribute is negative, then the parent run.shall
         * be lowered below the baseline of the surrounding text by the specified
         * number of half-points.
         * </p>
         * <p/>
         * If this element is not present, the default value is to leave the
         * formatting applied at previous level in the style hierarchy. If this
         * element is never applied in the style hierarchy, then the text shall not
         * be raised or lowered relative to the default baseline location for the
         * contents of this run.
         * </p>
         *
         * @param val
         */
        public void SetTextPosition(int val)
        {
            //Bigint bint=new Bigint(""+val);
            //CT_RPr pr = run.IsSetRPr() ? run.RPr : run.AddNewRPr();
            //CTSignedHpsMeasure position = pr.IsSetPosition() ? pr.Position : pr.AddNewPosition();
            //position.Val=(bint);
            throw new NotImplementedException();
        }

        /**
         * 
         */
        public void RemoveBreak()
        {
            // TODO
        }

        /**
         * Specifies that a break shall be placed at the current location in the run
         * content. 
         * A break is a special character which is used to override the
         * normal line breaking that would be performed based on the normal layout
         * of the document's contents. 
         * @see #AddCarriageReturn() 
         */
        public void AddBreak()
        {
            //run.AddNewBr();
            throw new NotImplementedException();
        }

        /**
         * Specifies that a break shall be placed at the current location in the run
         * content.
         * A break is a special character which is used to override the
         * normal line breaking that would be performed based on the normal layout
         * of the document's contents.
         * <p>
         * The behavior of this break character (the
         * location where text shall be restarted After this break) shall be
         * determined by its type values.
         * </p>
         * @see BreakType
         */
        public void AddBreak(BreakType type)
        {
            //CTBr br=run.AddNewBr();
            //br.Type=(STBrType.Enum.ForInt(type.Value));
            throw new NotImplementedException();
        }

        /**
         * Specifies that a break shall be placed at the current location in the run
         * content. A break is a special character which is used to override the
         * normal line breaking that would be performed based on the normal layout
         * of the document's contents.
         * <p>
         * The behavior of this break character (the
         * location where text shall be restarted After this break) shall be
         * determined by its type (in this case is BreakType.TEXT_WRAPPING as default) and clear attribute values.
         * </p>
         * @see BreakClear
         */
        public void AddBreak(BreakClear Clear)
        {
            //CTBr br=run.AddNewBr();
            //br.Type=(STBrType.Enum.ForInt(BreakType.TEXT_WRAPPING.Value));
            //br.Clear=(STBrClear.Enum.ForInt(Clear.Value));
            throw new NotImplementedException();
        }

        /**
         * Specifies that a carriage return shall be placed at the
         * current location in the run.content.
         * A carriage return is used to end the current line of text in
         * WordProcess.
         * The behavior of a carriage return in run.content shall be
         * identical to a break character with null type and clear attributes, which
         * shall end the current line and find the next available line on which to
         * continue.
         * The carriage return character forced the following text to be
         * restarted on the next available line in the document.
         */
        public void AddCarriageReturn()
        {
            //run.AddNewCr();
            throw new NotImplementedException();
        }

        public void RemoveCarriageReturn()
        {
            //TODO
        }

        /**
         * Adds a picture to the run. This method handles
         *  attaching the picture data to the overall file.
         *  
         * @see NPOI.XWPF.UserModel.Document#PICTURE_TYPE_EMF
         * @see NPOI.XWPF.UserModel.Document#PICTURE_TYPE_WMF
         * @see NPOI.XWPF.UserModel.Document#PICTURE_TYPE_PICT
         * @see NPOI.XWPF.UserModel.Document#PICTURE_TYPE_JPEG
         * @see NPOI.XWPF.UserModel.Document#PICTURE_TYPE_PNG
         * @see NPOI.XWPF.UserModel.Document#PICTURE_TYPE_DIB
         *  
         * @param pictureData The raw picture data
         * @param pictureType The type of the picture, eg {@link Document#PICTURE_TYPE_JPEG}
         * @throws IOException 
         * @throws NPOI.Openxml4j.exceptions.InvalidFormatException 
         * @throws IOException 
         */
        public XWPFPicture AddPicture(Stream pictureData, int pictureType, String filename, int width, int height)
        {
            //XWPFDocument doc = paragraph.document;

            //// Add the picture + relationship
            //String relationId = doc.AddPictureData(pictureData, pictureType);
            //XWPFPictureData picData = (XWPFPictureData) doc.GetRelationById(relationId);

            //// Create the Drawing entry for it
            //try {
            //    CTDrawing Drawing = run.AddNewDrawing();
            //    CTInline inline = Drawing.AddNewInline();

            //    // Do the fiddly namespace bits on the inline
            //    // (We need full control of what goes where and as what)
            //    String xml = 
            //        "<a:graphic xmlns:a=\"" + CTGraphicalObject.type.Name.NamespaceURI + "\">" +
            //        "<a:graphicData uri=\"" + CTGraphicalObject.type.Name.NamespaceURI + "\">" +
            //        "<pic:pic xmlns:pic=\"" + CT_Picture.type.Name.NamespaceURI + "\" />" +
            //        "</a:graphicData>" +
            //        "</a:graphic>";
            //    inline.Set(XmlToken.Factory.Parse(xml));

            //    // Setup the inline
            //    inline.DistT=(0);
            //    inline.DistR=(0);
            //    inline.DistB=(0);
            //    inline.DistL=(0);

            //    CTNonVisualDrawingProps docPr = inline.AddNewDocPr();
            //    long id = GetParagraph().document.DrawingIdManager.ReserveNew();
            //    docPr.Id=(id);
            //    /* This name is not visible in Word 2010 anywhere. */
            //    docPr.Name=("Drawing " + id);
            //    docPr.Descr=(filename);

            //    CTPositiveSize2D extent = inline.AddNewExtent();
            //    extent.Cx=(width);
            //    extent.Cy=(height);

            //    // Grab the picture object
            //    CTGraphicalObject graphic = inline.Graphic;
            //    CTGraphicalObjectData graphicData = graphic.GraphicData;
            //    CT_Picture pic = GetCT_Pictures(graphicData).Get(0);

            //    // Set it up
            //    CT_PictureNonVisual nvPicPr = pic.AddNewNvPicPr();

            //    CTNonVisualDrawingProps cNvPr = nvPicPr.AddNewCNvPr();
            //    /* use "0" for the id. See ECM-576, 20.2.2.3 */
            //    cNvPr.Id=(0L);
            //    /* This name is not visible in Word 2010 anywhere */
            //    cNvPr.Name=("Picture " + id);
            //    cNvPr.Descr=(filename);

            //    CTNonVisualPictureProperties cNvPicPr = nvPicPr.AddNewCNvPicPr();
            //    cNvPicPr.AddNewPicLocks().NoChangeAspect=(true);

            //    CTBlipFillProperties blipFill = pic.AddNewBlipFill();
            //    CTBlip blip = blipFill.AddNewBlip();
            //    blip.Embed=( picData.PackageRelationship.Id );
            //    blipFill.AddNewStretch().AddNewFillRect();

            //    CTShapeProperties spPr = pic.AddNewSpPr();
            //    CTTransform2D xfrm = spPr.AddNewXfrm();

            //    CTPoint2D off = xfrm.AddNewOff();
            //    off.X=(0);
            //    off.Y=(0);

            //    CTPositiveSize2D ext = xfrm.AddNewExt();
            //    ext.Cx=(width);
            //    ext.Cy=(height);

            //    CTPresetGeometry2D prstGeom = spPr.AddNewPrstGeom();
            //    prstGeom.Prst=(STShapeType.RECT);
            //    prstGeom.AddNewAvLst();

            //    // Finish up
            //    XWPFPicture xwpfPicture = new XWPFPicture(pic, this);
            //    pictures.Add(xwpfPicture);
            //    return xwpfPicture;
            //} catch(XmlException e) {
            //    throw new InvalidOperationException(e);
            //}
            throw new NotImplementedException();
        }

        /**
         * Returns the embedded pictures of the run. These
         *  are pictures which reference an external, 
         *  embedded picture image such as a .png or .jpg
         */
        public List<XWPFPicture> GetEmbeddedPictures()
        {
            return pictures;
        }

        /**
         * Add the xml:spaces="preserve" attribute if the string has leading or trailing white spaces
         *
         * @param xs    the string to check
         */
        static void preserveSpaces(CT_Text xs)
        {
            String text = xs.Value;
            if (text != null && (text.StartsWith(" ") || text.EndsWith(" "))) {
            //    XmlCursor c = xs.NewCursor();
            //    c.ToNextToken();
            //    c.InsertAttributeWithValue(new QName("http://www.w3.org/XML/1998/namespace", "space"), "preserve");
            //    c.Dispose();
            }
        }

        /**
         * Returns the string version of the text, with tabs and
         *  carriage returns in place of their xml equivalents.
         */
        public override String ToString()
        {
            StringBuilder text = new StringBuilder();

            //// Grab the text and tabs of the text run
            //// Do so in a way that preserves the ordering
            //XmlCursor c = run.NewCursor();
            //c.SelectPath("./*");
            //while (c.ToNextSelection()) {
            //    XmlObject o = c.Object;
            //    if (o is CTText) {
            //        String tagName = o.DomNode.NodeName;
            //        // Field Codes (w:instrText, defined in spec sec. 17.16.23)
            //        //  come up as instances of CTText, but we don't want them
            //        //  in the normal text output
            //        if (!"w:instrText".Equals(tagName)) {
            //            text.Append(((CTText) o).StringValue);
            //        }
            //    }

            //    if (o is CTPTab) {
            //        text.Append("\t");
            //    }
            //    if (o is CTBr) {
            //        text.Append("\n");
            //    }
            //    if (o is CTEmpty) {
            //        // Some inline text elements Get returned not as
            //        //  themselves, but as CTEmpty, owing to some odd
            //        //  defInitions around line 5642 of the XSDs
            //        // This bit works around it, and replicates the above
            //        //  rules for that case
            //        String tagName = o.DomNode.NodeName;
            //        if ("w:tab".Equals(tagName)) {
            //            text.Append("\t");
            //        }
            //        if ("w:br".Equals(tagName)) {
            //            text.Append("\n");
            //        }
            //        if ("w:cr".Equals(tagName)) {
            //            text.Append("\n");
            //        }
            //    }
            //}

            //c.Dispose();

            //// Any picture text?
            //if(pictureText != null && pictureText.Length() > 0) {
            //    text.Append("\n").Append(pictureText);
            //}

            //return text.ToString();
            throw new NotImplementedException();
        }
    }

}