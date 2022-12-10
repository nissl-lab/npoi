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
using NPOI.DDF;
using NPOI.SL.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace NPOI.HSLF.UserModel
{
	/**
 * Represents a Shape which is the elemental object that composes a drawing.
 *  This class is a wrapper around EscherSpContainer which holds all information
 *  about a shape in PowerPoint document.
 *  <p>
 *  When you add a shape, you usually specify the dimensions of the shape and the position
 *  of the upper'left corner of the bounding box for the shape relative to the upper'left
 *  corner of the page, worksheet, or slide. Distances in the drawing layer are measured
 *  in points (72 points = 1 inch).
 */
	public class HSLFShape: Shape<HSLFShape, HSLFTextParagraph>
	{
		/**
     * Either EscherSpContainer or EscheSpgrContainer record
     * which holds information about this shape.
     */
		private EscherContainerRecord _escherContainer;

		/**
		 * Parent of this shape.
		 * {@code null} for the topmost shapes.
		 */
		private  ShapeContainer<HSLFShape, HSLFTextParagraph> _parent;

		/**
		 * The {@code Sheet} this shape belongs to
		 */
		private HSLFSheet _sheet;

		/**
		 * Fill
		 */
		private HSLFFill _fill;

		/**
		 * Create a Shape object. This constructor is used when an existing Shape is read from a PowerPoint document.
		 *
		 * @param escherRecord       {@code EscherSpContainer} container which holds information about this shape
		 * @param parent             the parent of this Shape
		 */
		protected HSLFShape(EscherContainerRecord escherRecord, ShapeContainer<HSLFShape, HSLFTextParagraph> parent)
		{
			_escherContainer = escherRecord;
			_parent = parent;
		}

		/**
		 * Create and assign the lower level escher record to this shape
		 */
		protected EscherContainerRecord createSpContainer(bool isChild)
		{
			if (_escherContainer == null)
			{
				_escherContainer = new EscherContainerRecord();
				_escherContainer.SetOptions((short)15);
			}
			return _escherContainer;
		}

		/**
		 *  @return the parent of this shape
		 */
		//@Override
		public ShapeContainer<HSLFShape, HSLFTextParagraph> GetParent()
		{
			return _parent;
		}

		/**
		 * @return name of the shape.
		 */
		//@Override
		public String GetShapeName()
		{
			 EscherComplexProperty ep = getEscherProperty(getEscherOptRecord(), EscherProperties.GROUPSHAPE__SHAPENAME);
			if (ep != null)
			{
				 byte[] cd = ep.GetComplexData();
				return StringUtil.GetFromUnicodeLE0Terminated(cd, 0, cd.Length / 2);
			}
			else
			{
				return GetShapeType().nativeName + " " + GetShapeId();
			}
		}

		public ShapeType GetShapeType()
		{
			EscherSpRecord spRecord = GetEscherChild(EscherSpRecord.RECORD_ID);
			return ShapeType.forId(spRecord.GetShapeType(), false);
		}

		public void SetShapeType(ShapeType type)
		{
			EscherSpRecord spRecord = GetEscherChild(EscherSpRecord.RECORD_ID);
			spRecord.SetShapeType((short)type.nativeId);
			spRecord.SetVersion((short)0x2);
		}

		/**
		 * Returns the anchor (the bounding box rectangle) of this shape.
		 * All coordinates are expressed in points (72 dpi).
		 *
		 * @return the anchor of this shape
		 */
		//@Override
		//public Rectangle2D GetAnchor()
		//{
		//	EscherSpRecord spRecord = GetEscherChild(EscherSpRecord.RECORD_ID);
		//	int flags = spRecord.GetFlags();
		//	int x1, y1, x2, y2;
		//	EscherChildAnchorRecord childRec = GetEscherChild(EscherChildAnchorRecord.RECORD_ID);
		//	bool useChildRec = ((flags & EscherSpRecord.FLAG_CHILD) != 0);
		//	if (useChildRec && childRec != null)
		//	{
		//		x1 = childRec.GetDx1();
		//		y1 = childRec.GetDy1();
		//		x2 = childRec.GetDx2();
		//		y2 = childRec.GetDy2();
		//	}
		//	else
		//	{
		//		if (useChildRec)
		//		{
		//			LOG.atWarn().log("EscherSpRecord.FLAG_CHILD is set but EscherChildAnchorRecord was not found");
		//		}
		//		EscherClientAnchorRecord clientRec = getEscherChild(EscherClientAnchorRecord.RECORD_ID);
		//		if (clientRec == null)
		//		{
		//			throw new RecordFormatException("Could not read record 'CLIENT_ANCHOR' with record-id: " + EscherClientAnchorRecord.RECORD_ID);
		//		}
		//		x1 = clientRec.getCol1();
		//		y1 = clientRec.getFlag();
		//		x2 = clientRec.getDx1();
		//		y2 = clientRec.getRow1();
		//	}

		//	// TODO: find out where this -1 value comes from at #57820 (link to ms docs?)

		//	return new Rectangle2D.Double(
		//		(x1 == -1 ? -1 : Units.masterToPoints(x1)),
		//		(y1 == -1 ? -1 : Units.masterToPoints(y1)),
		//		(x2 == -1 ? -1 : Units.masterToPoints(x2 - x1)),
		//		(y2 == -1 ? -1 : Units.masterToPoints(y2 - y1))
		//	);
		//}

		/**
		 * Sets the anchor (the bounding box rectangle) of this shape.
		 * All coordinates should be expressed in points (72 dpi).
		 *
		 * @param anchor new anchor
		 */
		//public void setAnchor(Rectangle2D anchor)
		//{
		//	int x = Units.pointsToMaster(anchor.getX());
		//	int y = Units.pointsToMaster(anchor.getY());
		//	int w = Units.pointsToMaster(anchor.getWidth() + anchor.getX());
		//	int h = Units.pointsToMaster(anchor.getHeight() + anchor.getY());
		//	EscherSpRecord spRecord = getEscherChild(EscherSpRecord.RECORD_ID);
		//	int flags = spRecord.getFlags();
		//	if ((flags & EscherSpRecord.FLAG_CHILD) != 0)
		//	{
		//		EscherChildAnchorRecord rec = getEscherChild(EscherChildAnchorRecord.RECORD_ID);
		//		rec.setDx1(x);
		//		rec.setDy1(y);
		//		rec.setDx2(w);
		//		rec.setDy2(h);
		//	}
		//	else
		//	{
		//		EscherClientAnchorRecord rec = getEscherChild(EscherClientAnchorRecord.RECORD_ID);
		//		rec.setCol1((short)x);
		//		rec.setFlag((short)y);
		//		rec.setDx1((short)w);
		//		rec.setRow1((short)h);
		//	}

		//}

		/**
		 * Moves the top left corner of the shape to the specified point.
		 *
		 * @param x the x coordinate of the top left corner of the shape
		 * @param y the y coordinate of the top left corner of the shape
		 */
		public  void MoveTo(double x, double y)
		{
			// This convenience method should be implemented via setAnchor in subclasses
			// see HSLFGroupShape.setAnchor() for a reference
			Rectangle2D anchor = getAnchor();
			anchor.setRect(x, y, anchor.getWidth(), anchor.getHeight());
			setAnchor(anchor);
		}

		/**
		 * Helper method to return escher child by record ID
		 *
		 * @return escher record or {@code null} if not found.
		 */
		public static T GetEscherChild<T>(EscherContainerRecord owner, int recordId)where T: EscherRecord
		{
			return owner.GetChildById((short)recordId);
		}

		/**
		 * @since POI 3.14-Beta2
		 */
		public static T GetEscherChild<T>(EscherContainerRecord owner, EscherRecordTypes recordId)where T:EscherRecord
		{
			return GetEscherChild(owner, recordId.typeID);
		}

		public T GetEscherChild<T>(int recordId) where T: EscherRecord
		{
			return _escherContainer.GetChildById((short)recordId);
		}

		/**
		 * @since POI 3.14-Beta2
		 */
		public T GetEscherChild<T>(EscherRecordTypes recordId)where T: EscherRecord
		{
			return GetEscherChild(recordId.typeID);
		}

		/**
		 * Returns  escher property by id.
		 *
		 * @return escher property or {@code null} if not found.
		 *
		 * @deprecated use {@link #getEscherProperty(EscherPropertyTypes)} instead
		 */
		//@Deprecated
		//@Removal(version = "5.0.0")

		public static T GetEscherProperty<T>(AbstractEscherOptRecord opt, int propId)where T: EscherProperty
		{
			return (T)((opt == null) ? null : opt.Lookup(propId));
		}

		/**
		 * Returns  escher property by type.
		 *
		 * @return escher property or {@code null} if not found.
		 */
		public static T GetEscherProperty<T>(AbstractEscherOptRecord opt, EscherPropertyTypes type) where T: EscherProperty
		{
			return (opt == null) ? null : opt.Lookup(type);
		}

		/**
		 * Set an escher property for this shape.
		 *
		 * @param opt       The opt record to set the properties to.
		 * @param propId    The id of the property. One of the constants defined in EscherOptRecord.
		 * @param value     value of the property. If value = -1 then the property is removed.
		 *
		 * @deprecated use {@link #setEscherProperty(AbstractEscherOptRecord, EscherPropertyTypes, int)}
		 */
		//@Deprecated
		//@Removal(version = "5.0.0")

		public static void SetEscherProperty(AbstractEscherOptRecord opt, short propId, int value)
		{
			List<EscherProperty> props = opt.GetEscherProperties();
			for (IEnumerator<EscherProperty> iterator = props.GetEnumerator(); iterator.MoveNext();)
			{
				if (iterator.Current.GetPropertyNumber() == propId)
				{
					//iterator.remove;
					break;
				}
			}
			if (value != -1)
			{
				opt.AddEscherProperty(new EscherSimpleProperty(propId, value));
				opt.SortProperties();
			}
		}

		/**
		 * Set an escher property for this shape.
		 *
		 * @param opt       The opt record to set the properties to.
		 * @param propType  The type of the property.
		 * @param value     value of the property. If value = -1 then the property is removed.
		 */
		public static void SetEscherProperty(AbstractEscherOptRecord opt, EscherPropertyTypes propType, int value)
		{
			SetEscherProperty(opt, propType, false, value);
		}

		public static void SetEscherProperty(AbstractEscherOptRecord opt, EscherPropertyTypes propType, bool isBlipId, int value)
		{
			List<EscherProperty> props = opt.GetEscherProperties();
			for (IEnumerator<EscherProperty> iterator = props.GetEnumerator(); iterator.MoveNext();)
			{
				if (iterator.Current.GetPropertyNumber() == propType.propNumber)
				{
					//iterator.remove();
					break;
				}
			}
			if (value != -1)
			{
				opt.AddEscherProperty(new EscherSimpleProperty(propType, false, isBlipId, value));
				opt.SortProperties();
			}
		}



		/**
		 * Set an simple escher property for this shape.
		 *
		 * @param propId    The id of the property. One of the constants defined in EscherOptRecord.
		 * @param value     value of the property. If value = -1 then the property is removed.
		 *
		 * @deprecated use {@link #setEscherProperty(EscherPropertyTypes, int)}
		 */
		//@Deprecated
		//@Removal(version = "5.0.0")

		public void SetEscherProperty(short propId, int value)
		{
			AbstractEscherOptRecord opt = GetEscherOptRecord();
			SetEscherProperty(opt, propId, value);
		}

		/**
		 * Set an simple escher property for this shape.
		 *
		 * @param propType  The type of the property.
		 * @param value     value of the property. If value = -1 then the property is removed.
		 */
		public void SetEscherProperty(EscherPropertyTypes propType, int value)
		{
			AbstractEscherOptRecord opt = GetEscherOptRecord();
			SetEscherProperty(opt, propType, value);
		}

		/**
		 * Get the value of a simple escher property for this shape.
		 *
		 * @param propId    The id of the property. One of the constants defined in EscherOptRecord.
		 */
		public int GetEscherProperty(short propId)
		{
			AbstractEscherOptRecord opt = GetEscherOptRecord();
			EscherSimpleProperty prop = GetEscherProperty(opt, propId);
			return prop == null ? 0 : prop.GetPropertyValue();
		}

		/**
		 * Get the value of a simple escher property for this shape.
		 *
		 * @param propType    The type of the property. One of the constants defined in EscherOptRecord.
		 */
		public int GetEscherProperty(EscherPropertyTypes propType)
		{
			AbstractEscherOptRecord opt = GetEscherOptRecord();
			EscherSimpleProperty prop = GetEscherProperty(opt, propType);
			return prop == null ? 0 : prop.GetPropertyValue();
		}

		/**
		 * Get the value of a simple escher property for this shape.
		 *
		 * @param propId    The id of the property. One of the constants defined in EscherOptRecord.
		 *
		 * @deprecated use {@link #getEscherProperty(EscherPropertyTypes, int)} instead
		 */
		//@Deprecated
		//@Removal(version = "5.0.0")
		public int GetEscherProperty(short propId, int defaultValue)
		{
			AbstractEscherOptRecord opt = GetEscherOptRecord();
			EscherSimpleProperty prop = GetEscherProperty(opt, propId);
			return prop == null ? defaultValue : prop.GetPropertyValue();
		}

		/**
		 * Get the value of a simple escher property for this shape.
		 *
		 * @param type    The type of the property.
		 */
		public int GetEscherProperty(EscherPropertyTypes type, int defaultValue)
		{
			AbstractEscherOptRecord opt = GetEscherOptRecord();
			EscherSimpleProperty prop = GetEscherProperty(opt, type);
			return prop == null ? defaultValue : prop.GetPropertyValue();
		}

		/**
		 * @return  The shape container and its children that can represent this
		 *          shape.
		 */
		public EscherContainerRecord GetSpContainer()
		{
			return _escherContainer;
		}

		/**
		 * Event which fires when a shape is inserted in the sheet.
		 * In some cases we need to propagate changes to upper level containers.
		 * <br>
		 * Default implementation does nothing.
		 *
		 * @param sh - owning shape
		 */
		protected void AfterInsert(HSLFSheet sh)
		{
			if (_fill != null)
			{
				_fill.afterInsert(sh);
			}
		}

		/**
		 *  @return the {@code SlideShow} this shape belongs to
		 */
		//@Override
		public HSLFSheet GetSheet()
		{
			return _sheet;
		}

		/**
		 * Assign the {@code SlideShow} this shape belongs to
		 *
		 * @param sheet owner of this shape
		 */
		public void SetSheet(HSLFSheet sheet)
		{
			_sheet = sheet;
		}

		Color GetColor(short colorProperty, short opacityProperty)
		{
			 AbstractEscherOptRecord opt = GetEscherOptRecord();
			 EscherSimpleProperty colProp = GetEscherProperty(opt, colorProperty);
			 Color col;
			if (colProp == null)
			{
				col = Color.White;
			}
			else
			{
				EscherColorRef ecr = new EscherColorRef(colProp.GetPropertyValue());
				col = GetColor(ecr);
				if (col == null)
				{
					return null;
				}
			}

			double alpha = GetAlpha(opacityProperty);
			return new Color(); //Color(col.GetRed(), col.GetGreen(), col.GetBlue(), (int)(alpha * 255.0));
		}

		Color GetColor(EscherColorRef ecr)
		{
			bool fPaletteIndex = ecr.HasPaletteIndexFlag();
			bool fPaletteRGB = ecr.HasPaletteRGBFlag();
			bool fSystemRGB = ecr.HasSystemRGBFlag();
			bool fSchemeIndex = ecr.HasSchemeIndexFlag();
			bool fSysIndex = ecr.HasSysIndexFlag();

			int[] rgb = ecr.GetRGB();

			HSLFSheet sheet = GetSheet();
			if (fSchemeIndex && sheet != null)
			{
				//red is the index to the color scheme
				ColorSchemeAtom ca = sheet.GetColorScheme();
				int schemeColor = ca.GetColor(ecr.GetSchemeIndex());

				rgb[0] = (schemeColor >> 0) & 0xFF;
				rgb[1] = (schemeColor >> 8) & 0xFF;
				rgb[2] = (schemeColor >> 16) & 0xFF;
			}
			else if (fPaletteIndex)
			{
				//TODO
			}
			else if (fPaletteRGB)
			{
				//TODO
			}
			else if (fSystemRGB)
			{
				//TODO
			}
			else if (fSysIndex)
			{
				Color col = GetSysIndexColor(ecr);
				col = ApplySysIndexProcedure(ecr, col);
				return col;
			}

			return new Color(rgb[0], rgb[1], rgb[2]);
		}

		private Color GetSysIndexColor(EscherColorRef ecr)
		{
			SysIndexSource sis = ecr.GetSysIndexSource();
			if (sis == null)
			{
				int sysIdx = ecr.GetSysIndex();
				PresetColor pc = PresetColor.ValueOfNativeId(sysIdx);
				return (pc != null) ? pc.Color : null;
			}

			// TODO: check for recursive loops, when color getter also reference
			// a different color type
			switch (sis)
			{
				case FILL_COLOR:
					{
						return GetColor(EscherProperties.FILL__FILLCOLOR, EscherProperties.FILL__FILLOPACITY);
					}
				case LINE_OR_FILL_COLOR:
					{
						Color col = null;
						if (this is HSLFSimpleShape) {
							col = GetColor(EscherProperties.LINESTYLE__COLOR, EscherProperties.LINESTYLE__OPACITY);
						}
						if (col == null)
						{
							col = GetColor(EscherProperties.FILL__FILLCOLOR, EscherProperties.FILL__FILLOPACITY);
						}
						return col;
					}
				case LINE_COLOR:
					{
						if (this is HSLFSimpleShape) {
							return GetColor(EscherProperties.LINESTYLE__COLOR, EscherProperties.LINESTYLE__OPACITY);
						}
						break;
					}
				case SHADOW_COLOR:
					{
						if (this is HSLFSimpleShape) {
							return ((HSLFSimpleShape)this).GetShadowColor();
						}
						break;
					}
				case CURRENT_OR_LAST_COLOR:
					{
						// TODO ... read from graphics context???
						break;
					}
				case FILL_BACKGROUND_COLOR:
					{
						return getColor(EscherProperties.FILL__FILLBACKCOLOR, EscherProperties.FILL__FILLOPACITY);
					}
				case LINE_BACKGROUND_COLOR:
					{
						if (this is HSLFSimpleShape) {
							return ((HSLFSimpleShape)this).GetLineBackgroundColor();
						}
						break;
					}
				case FILL_OR_LINE_COLOR:
					{
						Color col = GetColor(EscherProperties.FILL__FILLCOLOR, EscherProperties.FILL__FILLOPACITY);
						if (col == null && this instanceof HSLFSimpleShape) {
							col = GetColor(EscherProperties.LINESTYLE__COLOR, EscherProperties.LINESTYLE__OPACITY);
						}
						return col;
					}
				default:
					break;
			}

			return null;
		}

		private Color ApplySysIndexProcedure(EscherColorRef ecr, Color col)
		{

			 SysIndexProcedure sip = ecr.GetSysIndexProcedure();
			if (col == null || sip == null)
			{
				return col;
			}

			switch (sip)
			{
				case DARKEN_COLOR:
					{
						// see java.awt.Color#darken()
						double FACTOR = (ecr.GetRGB()[2]) / 255;
						int r = (Math.Round(col.getRed() * FACTOR));
						int g = (Math.Round(col.getGreen() * FACTOR));
						int b = (Math.Round(col.getBlue() * FACTOR));
						return new Color(r, g, b);
					}
				case LIGHTEN_COLOR:
					{
						double FACTOR = (0xFF - ecr.getRGB()[2]) / 255.;

						int r = col.getRed();
						int g = col.getGreen();
						int b = col.getBlue();

						r = Math.toIntExact(Math.round(r + (0xFF - r) * FACTOR));
						g = Math.toIntExact(Math.round(g + (0xFF - g) * FACTOR));
						b = Math.toIntExact(Math.round(b + (0xFF - b) * FACTOR));

						return new Color(r, g, b);
					}
				default:
					// TODO ...
					break;
			}

			return col;
		}

		double GetAlpha(EscherPropertyTypes opacityProperty)
		{
			AbstractEscherOptRecord opt = GetEscherOptRecord();
			EscherSimpleProperty op = GetEscherProperty(opt, opacityProperty);
			int defaultOpacity = 0x00010000;
			int opacity = (op == null) ? defaultOpacity : op.GetPropertyValue();
			return Units.FixedPointToDouble(opacity);
		}

		Color toRGB(int val)
		{
			int a = (val >> 24) & 0xFF;
			int b = (val >> 16) & 0xFF;
			int g = (val >> 8) & 0xFF;
			int r = (val >> 0) & 0xFF;

			if (a == 0xFE)
			{
				// Color is an sRGB value specified by red, green, and blue fields.
			}
			else if (a == 0xFF)
			{
				// Color is undefined.
			}
			else
			{
				// index in the color scheme
				ColorSchemeAtom ca = GetSheet().GetColorScheme();
				int schemeColor = ca.GetColor(a);

				r = (schemeColor >> 0) & 0xFF;
				g = (schemeColor >> 8) & 0xFF;
				b = (schemeColor >> 16) & 0xFF;
			}
			return new Color(r, g, b);
		}

		//@Override
		public int GetShapeId()
		{
			EscherSpRecord spRecord = GetEscherChild(EscherSpRecord.RECORD_ID);
			return spRecord == null ? 0 : spRecord.GetShapeId();
		}

		/**
		 * Sets shape ID
		 *
		 * @param id of the shape
		 */
		public void SetShapeId(int id)
		{
			EscherSpRecord spRecord = GetEscherChild(EscherSpRecord.RECORD_ID);
			if (spRecord != null) spRecord.SetShapeId(id);
		}

		/**
		 * Fill properties of this shape
		 *
		 * @return fill properties of this shape
		 */
		public HSLFFill GetFill()
		{
			if (_fill == null)
			{
				_fill = new HSLFFill(this);
			}
			return _fill;
		}

		public FillStyle GetFillStyle()
		{
			return GetFill().GetFillStyle();
		}

		//@Override
		//public void Draw(Graphics2D graphics, Rectangle2D bounds)
		//{
		//	DrawFactory.getInstance(graphics).drawShape(graphics, this, bounds);
		//}

		public AbstractEscherOptRecord GetEscherOptRecord()
		{
			AbstractEscherOptRecord opt = GetEscherChild(EscherRecordTypes.OPT);
			if (opt == null)
			{
				opt = GetEscherChild(EscherRecordTypes.USER_DEFINED);
			}
			return opt;
		}

		public bool GetFlipHorizontal()
		{
			EscherSpRecord spRecord = GetEscherChild(EscherSpRecord.RECORD_ID);
			return (spRecord.GetFlags() & EscherSpRecord.FLAG_FLIPHORIZ) != 0;
		}

		public void SetFlipHorizontal(bool flip)
		{
			EscherSpRecord spRecord = GetEscherChild(EscherSpRecord.RECORD_ID);
			int flag = spRecord.GetFlags() | EscherSpRecord.FLAG_FLIPHORIZ;
			spRecord.SetFlags(flag);
		}

		public bool GetFlipVertical()
		{
			EscherSpRecord spRecord = GetEscherChild(EscherSpRecord.RECORD_ID);
			return (spRecord.GetFlags() & EscherSpRecord.FLAG_FLIPVERT) != 0;
		}

		public void SetFlipVertical(bool flip)
		{
			EscherSpRecord spRecord = GetEscherChild(EscherSpRecord.RECORD_ID);
			int flag = spRecord.GetFlags() | EscherSpRecord.FLAG_FLIPVERT;
			spRecord.SetFlags(flag);
		}

		public double GetRotation()
		{
			int rot = GetEscherProperty(EscherPropertyTypes.TRANSFORM__ROTATION);
			return Units.FixedPointToDouble(rot);
		}

		public void SetRotation(double theta)
		{
			int rot = Units.DoubleToFixedPoint(theta % 360.0);
			SetEscherProperty(EscherPropertyTypes.TRANSFORM__ROTATION, rot);
		}

		public bool IsPlaceholder()
		{
			return false;
		}

		/**
		 *  Find a record in the underlying EscherClientDataRecord
		 *
		 * @param recordType type of the record to search
		 */
		//@SuppressWarnings("unchecked")

		public T GetClientDataRecord<T>(int recordType) where T : Record.Record
		{

			List <Record.Record > records = GetClientRecords();
			if (records != null) foreach (Record.Record r in records)
				{
					if (r.GetRecordType() == recordType)
					{
						return (T)r;
					}
				}
			return null;
		}

		/**
		 * Search for EscherClientDataRecord, if found, convert its contents into an array of HSLF records
		 *
		 * @return an array of HSLF records contained in the shape's EscherClientDataRecord or {@code null}
		 */
		protected List<Record.Record> GetClientRecords()
		{
			HSLFEscherClientDataRecord clientData = GetClientData(false);
			return (clientData == null) ? null : clientData.GetHSLFChildRecords();
		}

		/**
		 * Create a new HSLF-specific EscherClientDataRecord
		 *
		 * @param create if true, create the missing record
		 * @return the client record or null if it was missing and create wasn't activated
		 */
		protected HSLFEscherClientDataRecord GetClientData(bool create)
		{
			HSLFEscherClientDataRecord clientData = GetEscherChild(EscherClientDataRecord.RECORD_ID);
			if (clientData == null && create)
			{
				clientData = new HSLFEscherClientDataRecord();
				clientData.SetOptions((short)15);
				clientData.SetRecordId(EscherClientDataRecord.RECORD_ID);
				GetSpContainer().AddChildBefore(clientData, EscherTextboxRecord.RECORD_ID);
			}
			return clientData;
		}
	}
}
