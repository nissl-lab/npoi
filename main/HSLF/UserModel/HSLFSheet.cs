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
using NPOI.SL.UserModel;
using NPOI.HSLF.Record;
using System.Collections.Generic;
using NPOI.HSLF.Exceptions;
using System;

namespace NPOI.HSLF.UserModel
{
	/**
 * This class defines the common format of "Sheets" in a powerpoint
 * document. Such sheets could be Slides, Notes, Master etc
 */
	public abstract class HSLFSheet : HSLFShapeContainer, Sheet<HSLFShape, HSLFTextParagraph>
	{
		/**
     * The {@code SlideShow} we belong to
     */
		private HSLFSlideShow _slideShow;

		/**
		 * Sheet background
		 */
		private HSLFBackground _background;

		/**
		 * Record container that holds sheet data.
		 * For slides it is org.apache.poi.hslf.record.Slide,
		 * for notes it is org.apache.poi.hslf.record.Notes,
		 * for slide masters it is org.apache.poi.hslf.record.SlideMaster, etc.
		 */
		private SheetContainer _container;

		private int _sheetNo;

		public HSLFSheet(SheetContainer container, int sheetNo)
		{
			_container = container;
			_sheetNo = sheetNo;
		}

		/**
		 * Returns an array of all the TextRuns in the sheet.
		 */
		public abstract List<List<HSLFTextParagraph>> GetTextParagraphs();

		/**
		 * Returns the (internal, RefID based) sheet number, as used
		 * to in PersistPtr stuff.
		 */
		public int _getSheetRefId()
		{
			return _container.GetSheetId();
		}

		/**
		 * Returns the (internal, SlideIdentifier based) sheet number, as used
		 * to reference this sheet from other records.
		 */
		public int _getSheetNumber()
		{
			return _sheetNo;
		}

		/**
		 * Fetch the PPDrawing from the underlying record
		 */
		public PPDrawing GetPPDrawing()
		{
			return _container.GetPPDrawing();
		}

		/**
		 * Fetch the SlideShow we're attached to
		 */
		//@Override
		public override HSLFSlideShow GetSlideShow()
		{
			return _slideShow;
		}

		/**
		 * Return record container for this sheet
		 */
		public SheetContainer GetSheetContainer()
		{
			return _container;
		}

		/**
		 * Set the SlideShow we're attached to.
		 * Also passes it on to our child text paragraphs
		 */
		//@Internal
		protected void SetSlideShow(HSLFSlideShow ss)
		{
			if (_slideShow != null)
			{
				throw new HSLFException("Can't change existing slideshow reference");
			}

			_slideShow = ss;
			List<List<HSLFTextParagraph>> trs = GetTextParagraphs();
			if (trs == null)
			{
				return;
			}
			foreach (List<HSLFTextParagraph> ltp in trs)
			{
				HSLFTextParagraph.SupplySheet(ltp, this);
				HSLFTextParagraph.ApplyHyperlinks(ltp);
			}
		}


		/**
		 * Returns all shapes contained in this Sheet
		 *
		 * @return all shapes contained in this Sheet (Slide or Notes)
		 */
		//@Override
		public List<HSLFShape> GetShapes()
		{
			PPDrawing ppdrawing = GetPPDrawing();

			EscherContainerRecord dg = ppdrawing.GetDgContainer();
			EscherContainerRecord spgr = null;

			foreach (EscherRecord rec in dg)
			{
				if (rec.GetRecordId() == EscherContainerRecord.SPGR_CONTAINER)
				{
					spgr = (EscherContainerRecord)rec;
					break;
				}
			}
			if (spgr == null)
			{
				throw new InvalidOperationException("spgr not found");
			}

			List<HSLFShape> shapeList = new List<HSLFShape>();
			bool isFirst = true;
			foreach (EscherRecord r in spgr)
			{
				if (isFirst)
				{
					// skip first item
					isFirst = false;
					continue;
				}

				EscherContainerRecord sp = (EscherContainerRecord)r;
				HSLFShape sh = HSLFShapeFactory.CreateShape(sp, null);
				sh.SetSheet(this);

				if (sh is HSLFSimpleShape)
				{
					HSLFHyperlink link = HSLFHyperlink.Find(sh);
					if (link != null)
					{
						((HSLFSimpleShape)sh).SetHyperlink(link);
					}
				}

				shapeList.Add(sh);
			}

			return shapeList;
		}

		/**
		 * Add a new Shape to this Slide
		 *
		 * @param shape - the Shape to add
		 */
		//@Override
		public void AddShape(HSLFShape shape)
		{
			PPDrawing ppdrawing = GetPPDrawing();

			EscherContainerRecord dgContainer = ppdrawing.GetDgContainer();
			EscherContainerRecord spgr = HSLFShape.GetEscherChild(dgContainer, EscherContainerRecord.SPGR_CONTAINER);
			spgr.addChildRecord(shape.getSpContainer());

			shape.setSheet(this);
			shape.setShapeId(AllocateShapeId());
			shape.afterInsert(this);
		}

		/**
		 * Allocates new shape id for the new drawing group id.
		 *
		 * @return a new shape id.
		 */
		public int AllocateShapeId()
		{
			EscherDggRecord dgg = _slideShow.GetDocumentRecord().GetPPDrawingGroup().GetEscherDggRecord();
			EscherDgRecord dg = _container.GetPPDrawing().GetEscherDgRecord();
			return dgg.AllocateShapeId(dg, false);
		}

		/**
		 * Removes the specified shape from this sheet.
		 *
		 * @param shape shape to be removed from this sheet, if present.
		 * @return {@code true} if the shape was deleted.
		 */
		//@Override
		public bool RemoveShape(HSLFShape shape)
		{
			PPDrawing ppdrawing = GetPPDrawing();

			EscherContainerRecord dg = ppdrawing.GetDgContainer();
			EscherContainerRecord spgr = dg.GetChildById(EscherContainerRecord.SPGR_CONTAINER);
			if (spgr == null)
			{
				return false;
			}

			return spgr.RemoveChildRecord(shape.GetSpContainer());
		}

		/**
		 * Called by SlideShow ater a new sheet is created
		 */
		public void OnCreate()
		{

		}

		/**
		 * Return the master sheet .
		 */
		//@Override
		public abstract HSLFMasterSheet GetMasterSheet();

		/**
		 * Color scheme for this sheet.
		 */
		public ColorSchemeAtom GetColorScheme()
		{
			return _container.GetColorScheme();
		}

		/**
		 * Returns the background shape for this sheet.
		 *
		 * @return the background shape for this sheet.
		 */
		//@Override
		public HSLFBackground GetBackground()
		{
			if (_background == null)
			{
				PPDrawing ppdrawing = GetPPDrawing();

				EscherContainerRecord dg = ppdrawing.GetDgContainer();
				EscherContainerRecord spContainer = dg.GetChildById(EscherContainerRecord.SP_CONTAINER);
				_background = new HSLFBackground(spContainer, null);
				_background.SetSheet(this);
			}
			return _background;
		}

		////@Override
		//public void draw(Graphics2D graphics)
		//{
		//	DrawFactory drawFact = DrawFactory.getInstance(graphics);
		//	Drawable draw = drawFact.getDrawable(this);
		//	draw.draw(graphics);
		//}

		/**
		 * Subclasses should call this method and update the array of text runs
		 * when a text shape is added
		 */
		protected void OnAddTextShape(HSLFTextShape shape)
		{
		}

		/**
		 * Return placeholder by text type
		 *
		 * @param type  type of text, See {@link org.apache.poi.hslf.record.TextHeaderAtom}
		 * @return  {@code TextShape} or {@code null}
		 */
		public HSLFTextShape GetPlaceholderByTextType(int type)
		{
			foreach (HSLFShape shape in GetShapes())
			{
				if (shape is HSLFTextShape)
				{
					HSLFTextShape tx = (HSLFTextShape)shape;
					if (tx.GetRunType() == type)
					{
						return tx;
					}
				}
			}
			return null;
		}

		/**
		 * Search placeholder by its type
		 *
		 * @param type  type of placeholder to search. See {@link org.apache.poi.hslf.record.OEPlaceholderAtom}
		 * @return  {@code SimpleShape} or {@code null}
		 */
		public HSLFSimpleShape GetPlaceholder(Placeholder type)
		{
			foreach (HSLFShape shape in GetShapes())
			{
				if (shape is HSLFSimpleShape)
				{
					HSLFSimpleShape ss = (HSLFSimpleShape)shape;
					if (type == ss.GetPlaceholder())
					{
						return ss;
					}
				}
			}
			return null;
		}

		/**
		 * Return programmable tag associated with this sheet, e.g. {@code ___PPT12}.
		 *
		 * @return programmable tag associated with this sheet.
		 */
		public string GetProgrammableTag()
		{
			string tag = null;
			RecordContainer progTags = (RecordContainer)
					GetSheetContainer().FindFirstOfType(
								RecordTypes.ProgTags.typeID
			);
			if (progTags != null)
			{
				RecordContainer progBinaryTag = (RecordContainer)
					progTags.FindFirstOfType(
							RecordTypes.ProgBinaryTag.typeID
				);
				if (progBinaryTag != null)
				{
					CString binaryTag = (CString)
						progBinaryTag.FindFirstOfType(
								RecordTypes.CString.typeID
					);
					if (binaryTag != null)
					{
						tag = binaryTag.GetText();
					}
				}
			}

			return tag;

		}

		//@Override
		public IEnumerator<HSLFShape> Iterator()
		{
			return GetShapes().GetEnumerator();
		}

		/**
		 * @since POI 5.2.0
		 */
		//@Override
		public Spliterator<HSLFShape> Spliterator()
		{
			return GetShapes().Spliterator();
		}

		/**
		 * @return whether shapes on the master sheet should be shown. By default master graphics is turned off.
		 * Sheets that support the notion of master (slide, slideLayout) should override it and
		 * check this setting
		 */
		//@Override
		public bool GetFollowMasterGraphics()
		{
			return false;
		}


		//@Override
		public HSLFTextBox CreateTextBox()
		{
			HSLFTextBox s = new HSLFTextBox();
			s.SetHorizontalCentered(true);
			s.SetAnchor(new Rectangle2D.Double(0, 0, 100, 100));
			AddShape(s);
			return s;
		}

		//@Override
		public HSLFAutoShape CreateAutoShape()
		{
			HSLFAutoShape s = new HSLFAutoShape(ShapeType.RECT);
			s.SetHorizontalCentered(true);
			s.SetAnchor(new Rectangle2D.Double(0, 0, 100, 100));
			AddShape(s);
			return s;
		}

		//@Override
		public HSLFFreeformShape CreateFreeform()
		{
			HSLFFreeformShape s = new HSLFFreeformShape();
			s.SetHorizontalCentered(true);
			s.SetAnchor(new Rectangle2D.Double(0, 0, 100, 100));
			AddShape(s);
			return s;
		}

		//@Override
		public HSLFConnectorShape CreateConnector()
		{
			HSLFConnectorShape s = new HSLFConnectorShape();
			s.SetAnchor(new Rectangle2D.Double(0, 0, 100, 100));
			AddShape(s);
			return s;
		}

		//@Override
		public HSLFGroupShape CreateGroup()
		{
			HSLFGroupShape s = new HSLFGroupShape();
			s.SetAnchor(new Rectangle2D.Double(0, 0, 100, 100));
			AddShape(s);
			return s;
		}

		//@Override
		public HSLFPictureShape CreatePicture(PictureData pictureData)
		{
			if (!(pictureData is HSLFPictureData))
			{
				throw new InvalidOperationException("pictureData needs to be of type HSLFPictureData");
			}
			HSLFPictureShape s = new HSLFPictureShape((HSLFPictureData)pictureData);
			s.SetAnchor(new Rectangle2D.Double(0, 0, 100, 100));
			AddShape(s);
			return s;
		}

		//@Override
		public HSLFTable CreateTable(int numRows, int numCols)
		{
			if (numRows < 1 || numCols < 1)
			{
				throw new InvalidOperationException("numRows and numCols must be greater than 0");
			}
			HSLFTable s = new HSLFTable(numRows, numCols);
			// anchor is set in constructor based on numRows/numCols
			AddShape(s);
			return s;
		}

		//@Override
		public HSLFObjectShape CreateOleShape(PictureData pictureData)
		{
			if (!(pictureData is HSLFPictureData))
			{
				throw new InvalidOperationException("pictureData needs to be of type HSLFPictureData");
			}
			HSLFObjectShape s = new HSLFObjectShape((HSLFPictureData)pictureData);
			s.SetAnchor(new Rectangle2D.Double(0, 0, 100, 100));
			AddShape(s);
			return s;
		}

		/**
		 * Header / Footer settings for this slide.
		 *
		 * @return Header / Footer settings for this slide
		 */
		public HeadersFooters GetHeadersFooters()
		{
			return new HeadersFooters(this, HeadersFootersContainer.SlideHeadersFootersContainer);
		}


		//@Override
		public HSLFPlaceholderDetails GetPlaceholderDetails(Placeholder placeholder)
		{
			HSLFSimpleShape ph = GetPlaceholder(placeholder);
			return (ph == null) ? null : new HSLFShapePlaceholderDetails(ph);
		}
	}
}