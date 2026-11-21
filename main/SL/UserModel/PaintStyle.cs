using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SL.UserModel
{
    /// <summary>
    /// The PaintStyle can be modified by secondary sources, e.g. the attributes in the preset shapes.
    /// These modifications need to be taken into account when the final color is determined
    /// </summary>
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

    public enum GradientType { linear, circular, rectangular, shape }
    public interface PaintStyle
    {

    }
    public interface SolidPaint : PaintStyle
    { 
        ColorStyle SolidColor { get; }
    }
    public interface GradientPaint : PaintStyle
    {
        double GradientAngle { get; }
        ColorStyle[] GradientColors { get; }
        float[] GradientFractions { get; }
        bool IsRotatedWithShape { get; }
        GradientType GradientType { get; }

        Insets2D GetFillToInsets();
    }
    public interface TexturePaint : PaintStyle
    { 
        
    }
}
