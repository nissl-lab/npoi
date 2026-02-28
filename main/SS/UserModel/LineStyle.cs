using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel
{
    public enum LineStyle:int
    {
        None = -1,
        Solid = 0, // Solid (continuous) pen
        DashSys = 1, // PS_DASH system   dash style
        DotSys = 2, // PS_DOT system   dash style
        DashDotSys = 3, // PS_DASHDOT system dash style
        DashDotDotSys = 4, // PS_DASHDOTDOT system dash style
        DotGel = 5, // square dot style
        DashGel = 6, // dash style
        LongDashGel = 7, // long dash style
        DashDotGel = 8, // dash short dash
        LongDashDotGel = 9, // long dash short dash
        LongDashDotDotGel = 10, // long dash short dash short dash
    }

    // End Line Cap
    public enum LineEndingCapType : int
    {
        None,
        Round,      // Rounded ends. Semi-circle protrudes by half line width.
        Square,     // Square protrudes by half line width.
        Flat,       // Line ends at end point.
    }

    // Compound Line Type
    public enum CompoundLineType : int
    {
        None,
        SingleLine,     // Single line: one normal width
        DoubleLines,    // Double lines of equal width
        ThickThin,      // Double lines: one thick, one thin
        ThinThick,      // Double lines: one thin, one thick
        TripleLines     // Three lines: thin, thick, thin
	}
}