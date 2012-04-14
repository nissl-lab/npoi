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
namespace NPOI.XWPF.extractor
{
    using System;




using org.apache.poi;
using org.apache.poi;
using org.apache.poi;
using NPOI.Openxml4j.exceptions;
using NPOI.Openxml4j.opc;
using NPOI.XWPF.Model;
using NPOI.XWPF.Model;
using NPOI.XWPF.UserModel;
using NPOI.XWPF.UserModel;
using NPOI.XWPF.UserModel;
using NPOI.XWPF.UserModel;
using NPOI.XWPF.UserModel;
using NPOI.XWPF.UserModel;
using NPOI.XWPF.UserModel;
using org.apache.xmlbeans;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;

/**
 * Helper class to extract text from an OOXML Word file
 */
public class XWPFWordExtractor : POIXMLTextExtractor {
   public static XWPFRelation[] SUPPORTED_TYPES = new XWPFRelation[] {
      XWPFRelation.DOCUMENT, XWPFRelation.TEMPLATE,
      XWPFRelation.MACRO_DOCUMENT, 
      XWPFRelation.MACRO_TEMPLATE_DOCUMENT
   };
   
	private XWPFDocument document;
	private bool fetchHyperlinks = false;
	
	public XWPFWordExtractor(OPCPackage Container) throws XmlException, OpenXML4JException, IOException {
		this(new XWPFDocument(Container));
	}
	public XWPFWordExtractor(XWPFDocument document) {
		base(document);
		this.document = document;
	}

	/**
	 * Should we also fetch the hyperlinks, when fetching 
	 *  the text content? Default is to only output the
	 *  hyperlink label, and not the contents
	 */
	public void SetFetchHyperlinks(bool fetch) {
		fetchHyperlinks = fetch;
	}
	
	public static void main(String[] args)  {
		if(args.Length < 1) {
			System.err.Println("Use:");
			System.err.Println("  HXFWordExtractor <filename.docx>");
			System.Exit(1);
		}
		POIXMLTextExtractor extractor = 
			new XWPFWordExtractor(POIXMLDocument.OpenPackage(
					args[0]
			));
		System.out.Println(extractor.Text);
	}
	
	public String GetText() {
		StringBuilder text = new StringBuilder();
		XWPFHeaderFooterPolicy hfPolicy = document.HeaderFooterPolicy;

		// Start out with all headers
		extractHeaders(text, hfPolicy);
		
		// First up, all our paragraph based text
		Iterator<XWPFParagraph> i = document.ParagraphsIterator;
		while(i.HasNext()) {
			XWPFParagraph paragraph = i.Next();

			try {
				CTSectPr ctSectPr = null;
				if (paragraph.CTP.PPr!=null) {
					ctSectPr = paragraph.CTP.PPr.SectPr;
				}

				XWPFHeaderFooterPolicy headerFooterPolicy = null;

				if (ctSectPr!=null) {
					headerFooterPolicy = new XWPFHeaderFooterPolicy(document, ctSectPr);
					extractHeaders(text, headerFooterPolicy);
				}

				// Do the paragraph text
				foreach(XWPFRun run in paragraph.Runs) {
				   text.Append(Run.ToString());
				   if(run is XWPFHyperlinkRun && fetchHyperlinks) {
				      XWPFHyperlink link = ((XWPFHyperlinkRun)Run).GetHyperlink(document);
				      if(link != null)
				         text.Append(" <" + link.URL + ">");
				   }
				}

				// Add comments
				XWPFCommentsDecorator decorator = new XWPFCommentsDecorator(paragraph, null);
				text.Append(decorator.CommentText).Append('\n');
				
				// Do endnotes and footnotes
				String footnameText = paragraph.FootnoteText;
			   if(footnameText != null && footnameText.Length() > 0) {
			      text.Append(footnameText + "\n");
			   }

				if (ctSectPr!=null) {
					extractFooters(text, headerFooterPolicy);
				}
			} catch (IOException e) {
				throw new POIXMLException(e);
			} catch (XmlException e) {
				throw new POIXMLException(e);
			}
		}

		// Then our table based text
		Iterator<XWPFTable> j = document.TablesIterator;
		while(j.HasNext()) {
			text.Append(j.Next().Text).Append('\n');
		}
		
		// Finish up with all the footers
		extractFooters(text, hfPolicy);
		
		return text.ToString();
	}

	private void extractFooters(StringBuilder text, XWPFHeaderFooterPolicy hfPolicy) {
		if(hfPolicy.FirstPageFooter != null) {
			text.Append( hfPolicy.FirstPageFooter.Text );
		}
		if(hfPolicy.EvenPageFooter != null) {
			text.Append( hfPolicy.EvenPageFooter.Text );
		}
		if(hfPolicy.DefaultFooter != null) {
			text.Append( hfPolicy.DefaultFooter.Text );
		}
	}

	private void extractHeaders(StringBuilder text, XWPFHeaderFooterPolicy hfPolicy) {
		if(hfPolicy.FirstPageHeader != null) {
			text.Append( hfPolicy.FirstPageHeader.Text );
		}
		if(hfPolicy.EvenPageHeader != null) {
			text.Append( hfPolicy.EvenPageHeader.Text );
		}
		if(hfPolicy.DefaultHeader != null) {
			text.Append( hfPolicy.DefaultHeader.Text );
		}
	}
}

