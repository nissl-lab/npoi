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
using NPOI.POIFS.Properties;
using NPOI.SL.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.UserModel
{
	public abstract class HSLFSimpleShape: HSLFShape, SimpleShape<HSLFShape, HSLFTextParagraph>
	{
		//private static Logger LOG = LogManager.getLogger(HSLFSimpleShape.class);

    public static double DEFAULT_LINE_WIDTH = 0.75;

		protected static short[] ADJUST_VALUES = {
            EscherProperties.GEOMETRY__ADJUSTVALUE,
            EscherProperties.GEOMETRY__ADJUST2VALUE,
            EscherProperties.GEOMETRY__ADJUST3VALUE,
            EscherProperties.GEOMETRY__ADJUST4VALUE,
            EscherProperties.GEOMETRY__ADJUST5VALUE,
            EscherProperties.GEOMETRY__ADJUST6VALUE,
            EscherProperties.GEOMETRY__ADJUST7VALUE,
            EscherProperties.GEOMETRY__ADJUST8VALUE,
            EscherProperties.GEOMETRY__ADJUST9VALUE,
            EscherProperties.GEOMETRY__ADJUST10VALUE
	};

	/**
     * Hyperlink
     */
	protected HSLFHyperlink _hyperlink;

	/**
     * Create a SimpleShape object and initialize it from the supplied Record container.
     *
     * @param escherRecord    <code>EscherSpContainer</code> container which holds information about this shape
     * @param parent    the parent of the shape
     */
	protected HSLFSimpleShape(EscherContainerRecord escherRecord, ShapeContainer<HSLFShape, HSLFTextParagraph> parent)
		:base(escherRecord, parent)
	{
		
	}

	/**
     * Create a new Shape
     *
     * @param isChild   <code>true</code> if the Line is inside a group, <code>false</code> otherwise
     * @return the record container which holds this shape
     */
	////@Override
	protected EscherContainerRecord createSpContainer(boolean isChild)
	{
		EscherContainerRecord ecr = super.createSpContainer(isChild);
		ecr.setRecordId(EscherContainerRecord.SP_CONTAINER);

		EscherSpRecord sp = new EscherSpRecord();
		int flags = EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE;
		if (isChild)
		{
			flags |= EscherSpRecord.FLAG_CHILD;
		}
		sp.setFlags(flags);
		ecr.addChildRecord(sp);

		AbstractEscherOptRecord opt = new EscherOptRecord();
		opt.setRecordId(EscherOptRecord.RECORD_ID);
		ecr.addChildRecord(opt);

		EscherRecord anchor;
		if (isChild)
		{
			anchor = new EscherChildAnchorRecord();
		}
		else
		{
			anchor = new EscherClientAnchorRecord();

			//hack. internal variable EscherClientAnchorRecord.shortRecord can be
			//initialized only in fillFields(). We need to set shortRecord=false;
			byte[] header = new byte[16];
			LittleEndian.putUShort(header, 0, 0);
			LittleEndian.putUShort(header, 2, 0);
			LittleEndian.putInt(header, 4, 8);
			anchor.fillFields(header, 0, null);
		}
		ecr.addChildRecord(anchor);

		return ecr;
	}

	/**
     *  Returns width of the line in in points
     */
	public double getLineWidth()
	{
		AbstractEscherOptRecord opt = getEscherOptRecord();
		EscherSimpleProperty prop = getEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINEWIDTH);
		return (prop == null) ? DEFAULT_LINE_WIDTH : Units.toPoints(prop.getPropertyValue());
	}

	/**
     *  Sets the width of line in in points
     *  @param width  the width of line in in points
     */
	public void setLineWidth(double width)
	{
		AbstractEscherOptRecord opt = getEscherOptRecord();
		setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINEWIDTH, Units.toEMU(width));
	}

	/**
     * Sets the color of line
     *
     * @param color new color of the line
     */
	public void setLineColor(Color color)
	{
		AbstractEscherOptRecord opt = getEscherOptRecord();
		if (color == null)
		{
			setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__NOLINEDRAWDASH, 0x80000);
		}
		else
		{
			int rgb = new Color(color.getBlue(), color.getGreen(), color.getRed(), 0).getRGB();
			setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__COLOR, rgb);
			setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__NOLINEDRAWDASH, 0x180018);
		}
	}

	/**
     * @return color of the line. If color is not set returns {@code null}
     */
	public Color getLineColor()
	{
		AbstractEscherOptRecord opt = getEscherOptRecord();

		EscherSimpleProperty p = getEscherProperty(opt, EscherPropertyTypes.LINESTYLE__NOLINEDRAWDASH);
		if (p != null && (p.getPropertyValue() & 0x8) == 0)
		{
			return null;
		}

		return getColor(EscherPropertyTypes.LINESTYLE__COLOR, EscherPropertyTypes.LINESTYLE__OPACITY);
	}

	/**
     * @return background color of the line. If color is not set returns {@code null}
     */
	public Color getLineBackgroundColor()
	{
		AbstractEscherOptRecord opt = getEscherOptRecord();

		EscherSimpleProperty p = getEscherProperty(opt, EscherPropertyTypes.LINESTYLE__NOLINEDRAWDASH);
		if (p != null && (p.getPropertyValue() & 0x8) == 0)
		{
			return null;
		}

		return getColor(EscherPropertyTypes.LINESTYLE__BACKCOLOR, EscherPropertyTypes.LINESTYLE__OPACITY);
	}

	/**
     * Sets the background color of line
     *
     * @param color new background color of the line
     */
	public void setLineBackgroundColor(Color color)
	{
		AbstractEscherOptRecord opt = getEscherOptRecord();
		if (color == null)
		{
			setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__NOLINEDRAWDASH, 0x80000);
			opt.removeEscherProperty(EscherPropertyTypes.LINESTYLE__BACKCOLOR);
		}
		else
		{
			int rgb = new Color(color.getBlue(), color.getGreen(), color.getRed(), 0).getRGB();
			setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__BACKCOLOR, rgb);
			setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__NOLINEDRAWDASH, 0x180018);
		}
	}

	/**
     * Gets line cap.
     *
     * @return cap of the line.
     */
	public LineCap getLineCap()
	{
		AbstractEscherOptRecord opt = getEscherOptRecord();
		EscherSimpleProperty prop = getEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINEENDCAPSTYLE);
		return (prop == null) ? LineCap.FLAT : LineCap.fromNativeId(prop.getPropertyValue());
	}

	/**
     * Sets line cap.
     *
     * @param pen new style of the line.
     */
	public void setLineCap(LineCap pen)
	{
		AbstractEscherOptRecord opt = getEscherOptRecord();
		setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINEENDCAPSTYLE, pen == LineCap.FLAT ? -1 : pen.nativeId);
	}

	/**
     * Gets line dashing.
     *
     * @return dashing of the line.
     */
	public LineDash getLineDash()
	{
		AbstractEscherOptRecord opt = getEscherOptRecord();
		EscherSimpleProperty prop = getEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINEDASHING);
		return (prop == null) ? LineDash.SOLID : LineDash.fromNativeId(prop.getPropertyValue());
	}

	/**
     * Sets line dashing.
     *
     * @param pen new style of the line.
     */
	public void setLineDash(LineDash pen)
	{
		AbstractEscherOptRecord opt = getEscherOptRecord();
		setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINEDASHING, pen == LineDash.SOLID ? -1 : pen.nativeId);
	}

	/**
     * Gets the line compound style
     *
     * @return the compound style of the line.
     */
	public LineCompound getLineCompound()
	{
		AbstractEscherOptRecord opt = getEscherOptRecord();
		EscherSimpleProperty prop = getEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINESTYLE);
		return (prop == null) ? LineCompound.SINGLE : LineCompound.fromNativeId(prop.getPropertyValue());
	}

	/**
     * Sets the line compound style
     *
     * @param style new compound style of the line.
     */
	public void setLineCompound(LineCompound style)
	{
		AbstractEscherOptRecord opt = getEscherOptRecord();
		setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINESTYLE, style == LineCompound.SINGLE ? -1 : style.nativeId);
	}

	/**
     * Returns line style. One of the constants defined in this class.
     *
     * @return style of the line.
     */
	//@Override
	public StrokeStyle getStrokeStyle()
	{
		return new StrokeStyle() {
			//@Override

			public PaintStyle getPaint()
		{
			return DrawPaint.createSolidPaint(HSLFSimpleShape.this.getLineColor());
		}

		//@Override

			public LineCap getLineCap()
		{
			return null;
		}

		//@Override

			public LineDash getLineDash()
		{
			return HSLFSimpleShape.this.getLineDash();
		}

		//@Override

			public LineCompound getLineCompound()
		{
			return HSLFSimpleShape.this.getLineCompound();
		}

		//@Override

			public double getLineWidth()
		{
			return HSLFSimpleShape.this.getLineWidth();
		}
	};
}

//@Override
	public Color getFillColor()
{
	return getFill().getForegroundColor();
}

//@Override
	public void setFillColor(Color color)
{
	getFill().setForegroundColor(color);
}

//@Override
	public Guide getAdjustValue(String name)
{
	if (name == null || !name.matches("adj([1-9]|10)?"))
	{
		LOG.atInfo().log("Adjust value '{}' not supported. Using default value.", name);
		return null;
	}

	name = name.replace("adj", "");
	if (name.isEmpty())
	{
		name = "1";
	}

	final int adjInt = Integer.parseInt(name);
	if (adjInt < 1 || adjInt > 10)
	{
		throw new HSLFException("invalid adjust value: " + adjInt);
	}


	EscherPropertyTypes escherProp = ADJUST_VALUES[adjInt - 1];

	int adjval = getEscherProperty(escherProp, -1);

	if (adjval == -1)
	{
		return null;
	}

	// Bug 59004
	// the adjust value are format dependent, we scale them up so they match the OOXML ones.
	// see https://social.msdn.microsoft.com/Forums/en-US/33e458e6-58df-48fe-9a10-e303ab08991d/preset-shapes-for-ppt?forum=os_binaryfile

	// usually we deal with length units and only very few degree units:
	boolean isDegreeUnit = false;
	switch (getShapeType())
	{
		case ARC:
		case BLOCK_ARC:
		case CHORD:
		case PIE:
			isDegreeUnit = (adjInt == 1 || adjInt == 2);
			break;
		case CIRCULAR_ARROW:
		case LEFT_CIRCULAR_ARROW:
		case LEFT_RIGHT_CIRCULAR_ARROW:
			isDegreeUnit = (adjInt == 2 || adjInt == 3 || adjInt == 4);
			break;
		case MATH_NOT_EQUAL:
			isDegreeUnit = (adjInt == 2);
			break;
	}

	Guide gd = new Guide();
	gd.setName(name);
	gd.setFmla("val " + Math.rint(adjval * (isDegreeUnit ? 65536. : 100000./ 21000.)));
	return gd;
}

//@Override
	public CustomGeometry getGeometry()
{
	PresetGeometries dict = PresetGeometries.getInstance();
	ShapeType st = getShapeType();
	String name = (st != null) ? st.getOoxmlName() : null;
	CustomGeometry geom = dict.get(name);
	if (geom == null)
	{
		if (name == null)
		{
			name = (st != null) ? st.toString() : "<unknown>";
		}
		LOG.atWarn().log("No preset shape definition for shapeType: {}", name);
	}

	return geom;
}


public double getShadowAngle()
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	EscherSimpleProperty prop = getEscherProperty(opt, EscherPropertyTypes.SHADOWSTYLE__OFFSETX);
	int offX = (prop == null) ? 0 : prop.getPropertyValue();
	prop = getEscherProperty(opt, EscherPropertyTypes.SHADOWSTYLE__OFFSETY);
	int offY = (prop == null) ? 0 : prop.getPropertyValue();
	return Math.toDegrees(Math.atan2(offY, offX));
}

public double getShadowDistance()
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	EscherSimpleProperty prop = getEscherProperty(opt, EscherPropertyTypes.SHADOWSTYLE__OFFSETX);
	int offX = (prop == null) ? 0 : prop.getPropertyValue();
	prop = getEscherProperty(opt, EscherPropertyTypes.SHADOWSTYLE__OFFSETY);
	int offY = (prop == null) ? 0 : prop.getPropertyValue();
	return Units.toPoints((long)Math.hypot(offX, offY));
}

/**
 * @return color of the line. If color is not set returns <code>java.awt.Color.black</code>
 */
public Color getShadowColor()
{
	Color clr = getColor(EscherPropertyTypes.SHADOWSTYLE__COLOR, EscherPropertyTypes.SHADOWSTYLE__OPACITY);
	return clr == null ? Color.black : clr;
}

//@Override
	public Shadow<HSLFShape, HSLFTextParagraph> getShadow()
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	if (opt == null)
	{
		return null;
	}
	EscherProperty shadowType = opt.lookup(EscherPropertyTypes.SHADOWSTYLE__TYPE);
	if (shadowType == null)
	{
		return null;
	}

	return new Shadow<HSLFShape, HSLFTextParagraph>(){
			//@Override

			public SimpleShape<HSLFShape, HSLFTextParagraph> getShadowParent()
	{
		return HSLFSimpleShape.this;
	}

	//@Override

			public double getDistance()
	{
		return getShadowDistance();
	}

	//@Override

			public double getAngle()
	{
		return getShadowAngle();
	}

	//@Override

			public double getBlur()
	{
		// TODO Auto-generated method stub
		return 0;
	}

	//@Override

			public SolidPaint getFillStyle()
	{
		return DrawPaint.createSolidPaint(getShadowColor());
	}

};
    }

    public DecorationShape getLineHeadDecoration()
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	EscherSimpleProperty prop = getEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINESTARTARROWHEAD);
	return (prop == null) ? null : DecorationShape.fromNativeId(prop.getPropertyValue());
}

public void setLineHeadDecoration(DecorationShape decoShape)
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINESTARTARROWHEAD, decoShape == null ? -1 : decoShape.nativeId);
}

public DecorationSize getLineHeadWidth()
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	EscherSimpleProperty prop = getEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINESTARTARROWWIDTH);
	return (prop == null) ? null : DecorationSize.fromNativeId(prop.getPropertyValue());
}

public void setLineHeadWidth(DecorationSize decoSize)
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINESTARTARROWWIDTH, decoSize == null ? -1 : decoSize.nativeId);
}

public DecorationSize getLineHeadLength()
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	EscherSimpleProperty prop = getEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINESTARTARROWLENGTH);
	return (prop == null) ? null : DecorationSize.fromNativeId(prop.getPropertyValue());
}

public void setLineHeadLength(DecorationSize decoSize)
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINESTARTARROWLENGTH, decoSize == null ? -1 : decoSize.nativeId);
}

public DecorationShape getLineTailDecoration()
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	EscherSimpleProperty prop = getEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINEENDARROWHEAD);
	return (prop == null) ? null : DecorationShape.fromNativeId(prop.getPropertyValue());
}

public void setLineTailDecoration(DecorationShape decoShape)
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINEENDARROWHEAD, decoShape == null ? -1 : decoShape.nativeId);
}

public DecorationSize getLineTailWidth()
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	EscherSimpleProperty prop = getEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINEENDARROWWIDTH);
	return (prop == null) ? null : DecorationSize.fromNativeId(prop.getPropertyValue());
}

public void setLineTailWidth(DecorationSize decoSize)
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINEENDARROWWIDTH, decoSize == null ? -1 : decoSize.nativeId);
}

public DecorationSize getLineTailLength()
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	EscherSimpleProperty prop = getEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINEENDARROWLENGTH);
	return (prop == null) ? null : DecorationSize.fromNativeId(prop.getPropertyValue());
}

public void setLineTailLength(DecorationSize decoSize)
{
	AbstractEscherOptRecord opt = getEscherOptRecord();
	setEscherProperty(opt, EscherPropertyTypes.LINESTYLE__LINEENDARROWLENGTH, decoSize == null ? -1 : decoSize.nativeId);
}



//@Override
	public LineDecoration getLineDecoration()
{
	return new LineDecoration() {

			//@Override

			public DecorationShape getHeadShape()
	{
		return HSLFSimpleShape.this.getLineHeadDecoration();
	}

	//@Override

			public DecorationSize getHeadWidth()
	{
		return HSLFSimpleShape.this.getLineHeadWidth();
	}

	//@Override

			public DecorationSize getHeadLength()
	{
		return HSLFSimpleShape.this.getLineHeadLength();
	}

	//@Override

			public DecorationShape getTailShape()
	{
		return HSLFSimpleShape.this.getLineTailDecoration();
	}

	//@Override

			public DecorationSize getTailWidth()
	{
		return HSLFSimpleShape.this.getLineTailWidth();
	}

	//@Override

			public DecorationSize getTailLength()
	{
		return HSLFSimpleShape.this.getLineTailLength();
	}
};
    }

    //@Override
	public HSLFShapePlaceholderDetails getPlaceholderDetails()
{
	return new HSLFShapePlaceholderDetails(this);
}


//@Override
	public Placeholder getPlaceholder()
{
	return getPlaceholderDetails().getPlaceholder();
}

//@Override
	public void setPlaceholder(Placeholder placeholder)
{
	getPlaceholderDetails().setPlaceholder(placeholder);
}


//@Override
	public void setStrokeStyle(Object...styles)
{
	if (styles.length == 0)
	{
		// remove stroke
		setLineColor(null);
		return;
	}

	// TODO: handle PaintStyle
	for (Object st : styles) {
	if (st instanceof Number) {
		setLineWidth(((Number)st).doubleValue());
	} else if (st instanceof LineCap) {
		setLineCap((LineCap)st);
	} else if (st instanceof LineDash) {
		setLineDash((LineDash)st);
	} else if (st instanceof LineCompound) {
		setLineCompound((LineCompound)st);
	} else if (st instanceof Color) {
		setLineColor((Color)st);
	}
}
    }

    //@Override
	public HSLFHyperlink getHyperlink()
{
	return _hyperlink;
}

//@Override
	public HSLFHyperlink createHyperlink()
{
	if (_hyperlink == null)
	{
		_hyperlink = HSLFHyperlink.createHyperlink(this);
	}
	return _hyperlink;
}

/**
 * Sets the hyperlink - used when the document is parsed
 *
 * @param link the hyperlink
 */
protected void setHyperlink(HSLFHyperlink link)
{
	_hyperlink = link;
}

//@Override
	public boolean isPlaceholder()
{
	// currently we only identify TextShapes as placeholders
	return false;
}
	}
}
