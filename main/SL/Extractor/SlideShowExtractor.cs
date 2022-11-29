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

namespace NPOI.SL.Extractor
{
	using System;
	using System.Text;

	using NPOI.HSSF.UserModel;
	using NPOI.POIFS.FileSystem;
	using NPOI;
	using NPOI.SL.UserModel;
	using System.Collections.Generic;
	using NPOI.SS.UserModel;
	using System.ComponentModel;
	using NPOI.SS.Formula.Eval;
	using System.Linq;

	public class SlideShowExtractor<S, P> : POITextExtractor
		where S : Shape<S,P> 
		where P : TextParagraph<S,P,TextRun>
	{
		// placeholder text for slide numbers
	    private static readonly string SLIDE_NUMBER_PH = "‹#›";

		protected readonly SlideShow<S, P> slideshow;

		private bool slidesByDefault = true;
	    private bool notesByDefault;
	    private bool commentsByDefault;
	    private bool masterByDefault;

		private Predicate<Object> filter = o => true;

		private bool doCloseFilesystem = true;

		public SlideShowExtractor(SlideShow<S, P> slideshow)
		{
			this.slideshow = slideshow;
		}

		/**
	     * Returns opened document
	     *
	     * @return the opened document
	     */
	    //override
	    public SlideShow<S, P> getDocument()
		{
			return slideshow;
		}

		/**
	     * Should a call to getText() return slide text? Default is yes
	     */
	    public void setSlidesByDefault(bool slidesByDefault)
		{
			this.slidesByDefault = slidesByDefault;
		}

		/**
	     * Should a call to getText() return notes text? Default is no
	     */
	    public void setNotesByDefault(bool notesByDefault)
		{
			this.notesByDefault = notesByDefault;
		}
	
	    /**
	     * Should a call to getText() return comments text? Default is no
	     */
	    public void setCommentsByDefault(bool commentsByDefault)
		{
			this.commentsByDefault = commentsByDefault;
		}
	
	    /**
	     * Should a call to getText() return text from master? Default is no
	     */
	    public void setMasterByDefault(bool masterByDefault)
		{
			this.masterByDefault = masterByDefault;
		}

		//Override
	    public POITextExtractor getMetadataTextExtractor()
		{
			return slideshow.getMetadataTextExtractor();
		}
	
	    /**
	     * Fetches all the slide text from the slideshow, but not the notes, unless
	     * you've called setSlidesByDefault() and setNotesByDefault() to change this
	     */
//	    Override
	    public String getText()
		{
			StringBuilder sb = new StringBuilder();
			foreach (Slide< S, P > slide in this.slideshow.Slides)
			{
				getText(slide, sb.Append);
			}
			
         return sb.ToString();
		}
	
	    public string getText(Slide<S, P> slide)
		{
			StringBuilder sb = new StringBuilder();
			getText(slide, sb.Append);
			return sb.ToString();
		}

		private void getText(Slide<S, P> slide, Func<string, StringBuilder> consumer)
		{
			if (slidesByDefault)
			{
				printShapeText(slide, consumer);
			}
			
			if (masterByDefault)
			{
				MasterSheet<S, P> ms = slide.MasterSheet;
				printSlideMaster(ms, consumer);
				
             // only print slide layout, if it's a different instance
				MasterSheet<S, P> sl = slide.SlideLayout;
				if (sl != ms)
				{
					printSlideMaster(sl, consumer);
				}
			}
			
	         if (commentsByDefault)
			{
				printComments(slide, consumer);
			}
			
	         if (notesByDefault)
			{
				printNotes(slide, consumer);
			}
		}

		private void printShapeText(Slide<S, P> slide, Func<string, StringBuilder> consumer)
		{
			throw new NotImplementedException();
		}

		private void printSlideMaster(MasterSheet<S, P> master, Func<string, StringBuilder> consumer)
		{
			if (master == null)
			{
				return;
			}
			foreach (Shape< S, P > shape in master)
			{
				if (shape is TextShape<S,P>) 
				{
					TextShape<S, P> ts = (TextShape<S, P>)shape;
					string text = ts.getText();
					if (text == null || string.IsNullOrEmpty(text) || "*".Equals(text))
					{
						continue;
					}

					if (ts.isPlaceholder())
					{
						// don't bother about boiler plate text on master sheets
						//LOG.atInfo().log("Ignoring boiler plate (placeholder) text on slide master: {}", text);
						continue;
					}

					printTextParagraphs(ts.getTextParagraphs(), consumer);
				}
			}
		}

		private void printTextParagraphs(List<P> paras, Func<string, StringBuilder> consumer)
		{
			printTextParagraphs(paras, consumer, "\n");
		}

		private void printTextParagraphs(List<P> paras, Func<string, StringBuilder> consumer, string trailer)
		{
			printTextParagraphs(paras, consumer, trailer, SlideShowExtractor<S,P>.replaceTextCap);
		}

		private void printTextParagraphs(List<P> paras, Func<string, StringBuilder> consumer, string trailer, Func<TextRun, string> converter)
		{
			foreach (P p in paras)
			{
				foreach (TextRun r in p)
				{
					if (filter(r))
					{
						consumer(converter(r));
					}
				}
				if (!string.IsNullOrEmpty(trailer) && filter(trailer))
				{
					consumer(trailer);
				}
			}
		}

		private void printHeaderFooter(Sheet<S, P> sheet, Func<string, StringBuilder> consumer, Func<string, StringBuilder> footerCon)
		{
			Sheet<S, P> m = (sheet is Slide<S, P>) ? sheet.getMasterSheet() : sheet;
			var phH = Placeholder.PlaceholderEnum["HEADER"];
			var phF = Placeholder.PlaceholderEnum["FOOTER"];
			addSheetPlaceholderDatails(sheet, new Placeholder(phH[0], phH[1], phH[2], phH[3], phH[4]), consumer);
			addSheetPlaceholderDatails(sheet, new Placeholder(phF[0], phF[1], phF[2], phF[3], phF[4]), footerCon);

			if (!masterByDefault)
			{
				return;
			}

			// write header texts and determine footer text
			foreach (Shape<S, P> s in m)
			{
				if (!(s is TextShape<S, P>))
				{
					continue;
				}
				TextShape<S, P> ts = (TextShape<S, P>)s;
				PlaceholderDetails pd = ts.getPlaceholderDetails();
				if (pd == null || !pd.isVisible() || pd.getPlaceholder() == null)
				{
					continue;
				}
				switch (pd.getPlaceholder().nativeEnum)
				{
					case "HEADER":
						printTextParagraphs(ts.getTextParagraphs(), consumer);
						break;
					case "FOOTER":
						printTextParagraphs(ts.getTextParagraphs(), footerCon);
						break;
					case "SLIDE_NUMBER":
						printTextParagraphs(ts.getTextParagraphs(), footerCon, "\n", SlideShowExtractor<S,P>.replaceSlideNumber);
						break;
					case "DATETIME":
					// currently not supported
					default:
						break;
				}
			}
		}

		private void addSheetPlaceholderDatails(Sheet<S, P> sheet, Placeholder placeholder, Func<string, StringBuilder> consumer)
		{ 
			PlaceholderDetails headerPD = sheet.getPlaceholderDetails(placeholder);
			string headerStr = (headerPD != null) ? headerPD.getText() : null;
			if (headerStr != null && filter(headerPD))
			{
				consumer(headerStr);
			}
		}

		private void printShapeText(Sheet<S, P> sheet, Func<string, StringBuilder> consumer)
		{
			StringBuilder footer = new StringBuilder();
			printHeaderFooter(sheet, consumer, footer.Append);
			printShapeText((Sheet<S, P>)(ShapeContainer<S, P>)sheet, consumer);
			consumer(footer.ToString());
		}

		//@SuppressWarnings("unchecked")
		private void printShapeText(ShapeContainer<S, P> container, Func<string, StringBuilder> consumer)
		{
			foreach (Shape<S, P> shape in container)
			{
				if (shape is TextShape<S,P>) 
				{
					printTextParagraphs(((TextShape<S, P>)shape).getTextParagraphs(), consumer);
				} else if (shape is TableShape<S,P>) 
				{
					printShapeText((Slide<S, P>)(TableShape<S, P>)shape, consumer);
				} else if (shape is ShapeContainer<S,P>) 
				{
					printShapeText((ShapeContainer<S, P>)shape, consumer);
				}
			}
		}

		//@SuppressWarnings("Duplicates")
		private void printShapeText(TableShape<S, P> shape, Func<string, StringBuilder> consumer)
		{
			int nrows = shape.getNumberOfRows();
			int ncols = shape.getNumberOfColumns();
			for (int row = 0; row < nrows; row++)
			{
				String trailer = "";
				for (int col = 0; col < ncols; col++)
				{
					TableCell<S, P> cell = shape.getCell(row, col);
					//defensive null checks; don't know if they're necessary
					if (cell != null)
					{
						trailer = col < ncols - 1 ? "\t" : "\n";
						printTextParagraphs(cell.getTextParagraphs(), consumer, trailer);
					}
				}
				if (!trailer.Equals("\n") && filter("\n"))
				{
					consumer("\n");
				}
			}
		}

		private void printComments(Slide<S, P> slide, Func<string, StringBuilder> consumer)
		{
			slide.getComments().Select(c => consumer(c.getAuthor() + " - " + c.getText()));
		}

		private void printNotes(Slide<S, P> slide, Func<string, StringBuilder> consumer)
		{
			Notes<S, P> notes = slide.getNotes();
			if (notes == null)
			{
				return;
			}

			StringBuilder footer = new StringBuilder();
			printHeaderFooter(notes, consumer, footer.Append);
			printShapeText(notes, consumer);
			consumer(footer.ToString());
		}

		public List<ObjectShape<S, P>> getOLEShapes()
		{
			List<ObjectShape< S,P >> oleShapes = new List<ObjectShape<S, P>>();

			foreach (Slide< S, P > slide in slideshow.Slides)
			{
				addOLEShapes(oleShapes, slide);
			}

			return oleShapes;
		}

		//@SuppressWarnings("unchecked")
		private void addOLEShapes(List<ObjectShape<S, P>> oleShapes, ShapeContainer<S, P> container)
		{
			foreach (Shape<S, P> shape in container)
			{
				if (shape is ShapeContainer<S,P>) 
				{
					addOLEShapes(oleShapes, (ShapeContainer<S, P>)shape);
				} else if (shape is ObjectShape) 
				{
					oleShapes.add((ObjectShape<S, P>)shape);
				}
			}
		}

		private static string replaceSlideNumber(TextRun tr)
		{
			string raw = tr.getRawText();

			if (!raw.contains(SLIDE_NUMBER_PH))
			{
				return raw;
			}

			TextParagraph <S,P,TextRun> tp = tr.getParagraph();
			TextShape <S,P> ps = (tp != null) ? tp.getParentShape() : null;
			Sheet <S,P> sh = (ps != null) ? ps.getSheet() : null;
			string slideNr = (sh is Slide<S,P>) ? Int32.ToString(((Slide <S,P>)sh).getSlideNumber() + 1) : "";

			return raw.replace(SLIDE_NUMBER_PH, slideNr);
		}

		private static string replaceTextCap(TextRun tr)
		{
			TextParagraph<S,P,TextRun> tp = tr.getParagraph();
			TextShape<S,P> sh = (tp != null) ? tp.getParentShape() : null;
			Placeholder ph = (sh != null) ? sh.getPlaceholder() : null;

			// 0xB acts like cariage return in page titles and like blank in the others
			char sep = (
				ph.nativeEnum == "TITLE" ||
				ph.nativeEnum == "CENTERED_TITLE" ||
				ph.nativeEnum == "SUBTITLE"
			) ? '\n' : ' ';

			// PowerPoint seems to store files with \r as the line break
			// The messes things up on everything but a Mac, so translate them to \n
			string txt = tr.getRawText();
			txt = txt.Replace('\r', '\n');
			txt = txt.Replace((char)0x0B, sep);

			switch (tr.getTextCap())
			{
				case ALL:
					txt = txt.ToUpper(LocaleUtil.getUserLocale());
					break;
				case SMALL:
					txt = txt.ToLower(LocaleUtil.getUserLocale());
					break;
			}

			return txt;
		}

		/**
     * Extract the used codepoints for font embedding / subsetting. This method is not intended for public use.
     *
     * @param typeface the typeface/font family of the textruns to examine
     * @param italic use {@code true} for italic TextRuns, {@code false} for non-italic ones and
     *      {@code null} if it doesn't matter
     * @param bold use {@code true} for bold TextRuns, {@code false} for non-bold ones and
     *      {@code null} if it doesn't matter
     * @return a bitset with the marked/used codepoints
     */
		//@Internal
	public SparseBitSet getCodepointsInSparseBitSet(string typeface, bool italic, bool bold)
		{
			SparseBitSet glyphs = new SparseBitSet();

			Predicate<Object> filterOld = filter;
			try
			{
				filter = o => filterFonts(o, typeface, italic, bold);
				slideshow.getSlides().forEach(slide->
						getText(slide, s->s.codePoints().forEach(glyphs::set))
				);
			}
			finally
			{
				filter = filterOld;
			}
			return glyphs;
		}

		private static bool filterFonts(object o, string typeface, bool italic, bool bold)
		{
			if (!(o is TextRun)) {
				return false;
			}
			TextRun tr = (TextRun)o;
			return
				typeface.Equals(tr.getFontFamily(), StringComparison.InvariantCultureIgnoreCase) &&
				(italic == null || tr.isItalic() == italic) &&
				(bold == null || tr.isBold() == bold);
		}
		//@Override
		public void setCloseFilesystem(bool doCloseFilesystem)
		{
			this.doCloseFilesystem = doCloseFilesystem;
		}
		//@Override
		public bool isCloseFilesystem()
		{
			return doCloseFilesystem;
		}

		//@Override
		public SlideShow<S, P> getFilesystem()
		{
			return getDocument();
		}

		public override string Text => throw new NotImplementedException();

		public override POITextExtractor MetadataTextExtractor => throw new NotImplementedException();
	}
}
