using NPOI.SS.Formula.Eval;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SL.UserModel
{
	/**
	     * The PaintStyle can be modified by secondary sources, e.g. the attributes in the preset shapes.
	     * These modifications need to be taken into account when the final color is determined.
	     */
	public enum PaintModifier
	{
		/** don't use any paint/fill */
		NONE,
		/** use the paint/filling as-is */
		NORM,
		/** lighten the paint/filling */
		LIGHTEN,
		/** lighten (... a bit less) the paint/filling */
		LIGHTEN_LESS,
		/** darken the paint/filling */
		DARKEN,
		/** darken (... a bit less) the paint/filling */
		DARKEN_LESS
	}

	public enum FlipMode
	{
		/** not flipped/mirrored */
		NONE,
		/** flipped/mirrored/duplicated along the x axis */
		X,
		/** flipped/mirrored/duplicated along the y axis */
		Y,
		/** flipped/mirrored/duplicated along the x and y axis */
		XY
	}

	public enum TextureAlignment
	{
		BOTTOM,
		BOTTOM_LEFT,
		BOTTOM_RIGHT,
		CENTER,
		LEFT,
		RIGHT,
		TOP,
		TOP_LEFT,
		TOP_RIGHT
	}

	public interface ISolidPaint: IPaintStyle
	{
		ColorStyle SolidColor { get; }
	}
	public enum GradientType { linear, circular, rectangular, shape }
	public interface IGradientPaint : IPaintStyle
	{
		/**
		 * @return the angle of the gradient
		 */
		double GradientAngle { get; }

		ColorStyle[] GradientColors { get; }

		float[] GradientFractions { get; }

		bool IsRotatedWithShape();

		GradientType GradientType1 { get; }

		Insets2D FillToInsets { get; }
	}

	public interface ITexturePaint : IPaintStyle
	{
		/**
		 * @return the raw image stream
		 */
		InputStream ImageData { get; }

		/**
		 * @return the content type of the image data
		 */
		string ContentType { get; }

		/**
		 * @return the alpha mask in percents [0..100000]
		 */
		int Alpha { get; }

		/**
		 * @return {@code true}, if the rotation of the shape is also applied to the texture paint
		 */
		bool IsRotatedWithShape();

		/**
		 * @return the dimensions of the tiles in percent of the shape dimensions
		 * or {@code null} if no scaling is applied
		 */
		Dimension2D Scale { get; }

		/**
		 * @return the offset of the tiles in points or {@code null} if there's no offset
		 */
		Point2D Offset { get; }

		/**
		 * @return the flip/mirroring/duplication mode
		 */
		FlipMode FlipMode { get; }

		TextureAlignment Alignment { get; }

		/**
		 * Specifies the portion of the blip or image that is used for the fill.<p>
		 *
		 * Each edge of the image is defined by a percentage offset from the edge of the bounding box.
		 * A positive percentage specifies an inset and a negative percentage specifies an outset.<p>
		 *
		 * The percentage are ints based on 100000, so 100% = 100000.<p>
		 *
		 * So, for example, a left offset of 25% specifies that the left edge of the image is located
		 * to the right of the bounding box's left edge by 25% of the bounding box's width.
		 *
		 * @return the cropping insets of the source image
		 */
		Insets2D Insets { get; }

		/**
		 * The stretch specifies the edges of a fill rectangle.<p>
		 *
		 * Each edge of the fill rectangle is defined by a percentage offset from the corresponding edge
		 * of the picture's bounding box. A positive percentage specifies an inset and a negative percentage
		 * specifies an outset.<p>
		 *
		 * The percentage are ints based on 100000, so 100% = 100000.
		 *
		 * @return the stretching in the destination image
		 */
		Insets2D Stretch { get; }


		/**
		 * For pattern images, the duo tone defines the black/white pixel color replacement
		 */
		List<ColorStyle> DuoTone { get; }

		/**
		 * @return the shape this texture paint is applied to
		 */
		Shape<S, P> GetShape<S, P>() where S : Shape<S, P> where P : TextParagraph<S, P, object>;
	}

	public interface IPaintStyle
	{
		PaintModifier PaintModifier { get; }
		FlipMode FlipMode { get; }
		TextureAlignment TextureAlignment { get; }
	}
}
