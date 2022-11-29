using NPOI.Util;
using static NPOI.HSSF.Util.HSSFColor;
using System.Text;
using System;
using System.Collections.Generic;

namespace NPOI.SL.UserModel
{
	public class ShapeType
	{
		/** Preset-ID for XML-based shapes */
	    public int ooxmlId;
	    
	    /** Preset-ID for binary-based shapes */
	    public int nativeId;
	    
	    /** POI-specific name for the binary-based type */
	    public string nativeName;

		public static readonly Dictionary<string, (int OoxmlId, int NativeId, string NativeName)> ShapeTypeEnum = new Dictionary<string, (int ooxmlId, int nativeId, string nativeName)>
		{
			{ "NOT_PRIMITIVE",(-1, 0, "NotPrimitive") },
            { "LINE",(1, 20, "Line") },
            { "LINE_INV",(2, -1, null) },
            { "TRIANGLE",(3, 5, "IsocelesTriangle") },
            { "RT_TRIANGLE",(4, 6, "RightTriangle") },
            { "RECT",(5, 1, "Rectangle") },
            { "DIAMOND",(6, 4, "Diamond") },
            { "PARALLELOGRAM",(7, 7, "Parallelogram") },
            { "TRAPEZOID",(8, 8, "Trapezoid") },
            { "NON_ISOSCELES_TRAPEZOID",(9, -1, null) },
            { "PENTAGON",(10, 56, "Pentagon") },
            { "HEXAGON",(11, 9, "Hexagon") },
            { "HEPTAGON",(12, -1, null) },
            { "OCTAGON",(13, 10, "Octagon") },
            { "DECAGON",(14, -1, null) },
            { "DODECAGON",(15, -1, null) },
            { "STAR_4",(16, 187, "Star4") },
            { "STAR_5",(17, 12, "Star") }, // aka star in native
            { "STAR_6",(18, -1, null) },
            { "STAR_7",(19, -1, null) },
            { "STAR_8",(20, 58, "Star8") },
            { "STAR_10",(21, -1, null) },
            { "STAR_12",(22, -1, null) },
            { "STAR_16",(23, 59, "Star16") },
            { "SEAL",(23, 18, "Seal") }, // same as star_16, but twice in native
            { "STAR_24",(24, 92, "Star24") },
            { "STAR_32",(25, 60, "Star32") },
            { "ROUND_RECT",(26, 2, "RoundRectangle") },
            { "ROUND_1_RECT",(27, -1, null) },
            { "ROUND_2_SAME_RECT",(28, -1, null) },
            { "ROUND_2_DIAG_RECT",(29, -1, null) },
            { "SNIP_ROUND_RECT",(30, -1, null) },
            { "SNIP_1_RECT",(31, -1, null) },
            { "SNIP_2_SAME_RECT",(32, -1, null) },
            { "SNIP_2_DIAG_RECT",(33, -1, null) },
            { "PLAQUE",(34, 21, "Plaque") },
            { "ELLIPSE",(35, 3, "Ellipse") },
            { "TEARDROP",(36, -1, null) },
            { "HOME_PLATE",(37, 15, "HomePlate") },
            { "CHEVRON",(38, 55, "Chevron") },
            { "PIE_WEDGE",(39, -1, null) },
            { "PIE",(40, -1, null) },
            { "BLOCK_ARC",(41, 95, "BlockArc") },
            { "DONUT",(42, 23, "Donut") },
            { "NO_SMOKING",(43, 57, "NoSmoking") },
            { "RIGHT_ARROW",(44, 13, "Arrow") }, // aka arrow in native
            { "LEFT_ARROW",(45, 66, "LeftArrow") },
            { "UP_ARROW",(46, 68, "UpArrow") },
            { "DOWN_ARROW",(47, 67, "DownArrow") },
            { "STRIPED_RIGHT_ARROW",(48, 93, "StripedRightArrow") },
            { "NOTCHED_RIGHT_ARROW",(49, 94, "NotchedRightArrow") },
            { "BENT_UP_ARROW",(50, 90, "BentUpArrow") },
            { "LEFT_RIGHT_ARROW",(51, 69, "LeftRightArrow") },
            { "UP_DOWN_ARROW",(52, 70, "UpDownArrow") },
            { "LEFT_UP_ARROW",(53, 89, "LeftUpArrow") },
            { "LEFT_RIGHT_UP_ARROW",(54, 182, "LeftRightUpArrow") },
            { "QUAD_ARROW",(55, 76, "QuadArrow") },
            { "LEFT_ARROW_CALLOUT",(56, 77, "LeftArrowCallout") },
            { "RIGHT_ARROW_CALLOUT",(57, 78, "RightArrowCallout") },
            { "UP_ARROW_CALLOUT",(58, 79, "UpArrowCallout") },
            { "DOWN_ARROW_CALLOUT",(59, 80, "DownArrowCallout") },
            { "LEFT_RIGHT_ARROW_CALLOUT",(60, 81, "LeftRightArrowCallout") },
            { "UP_DOWN_ARROW_CALLOUT",(61, 82, "UpDownArrowCallout") },
            { "QUAD_ARROW_CALLOUT",(62, 83, "QuadArrowCallout") },
            { "BENT_ARROW",(63, 91, "BentArrow") },
            { "UTURN_ARROW",(64, 101, "UturnArrow") },
            { "CIRCULAR_ARROW",(65, 99, "CircularArrow") },
            { "LEFT_CIRCULAR_ARROW",(66, -1, null) },
            { "LEFT_RIGHT_CIRCULAR_ARROW",(67, -1, null) },
            { "CURVED_RIGHT_ARROW",(68, 102, "CurvedRightArrow") },
            { "CURVED_LEFT_ARROW",(69, 103, "CurvedLeftArrow") },
            { "CURVED_UP_ARROW",(70, 104, "CurvedUpArrow") },
            { "CURVED_DOWN_ARROW",(71, 105, "CurvedDownArrow") },
            { "SWOOSH_ARROW",(72, -1, null) },
            { "CUBE",(73, 16, "Cube") },
            { "CAN",(74, 22, "Can") },
            { "LIGHTNING_BOLT",(75, 73, "LightningBolt") },
            { "HEART",(76, 74, "Heart") },
            { "SUN",(77, 183, "Sun") },
            { "MOON",(78, 184, "Moon") },
            { "SMILEY_FACE",(79, 96, "SmileyFace") },
            { "IRREGULAR_SEAL_1",(80, 71, "IrregularSeal1") },
            { "IRREGULAR_SEAL_2",(81, 72, "IrregularSeal2") },
            { "FOLDED_CORNER",(82, 65, "FoldedCorner") },
            { "BEVEL",(83, 84, "Bevel") },
            { "FRAME",(84, 75, "PictureFrame") },
            { "HALF_FRAME",(85, -1, null) },
            { "CORNER",(86, -1, null) },
            { "DIAG_STRIPE",(87, -1, null) },
            { "CHORD",(88, -1, null) },
            { "ARC",(89, 19, "Arc") },
            { "LEFT_BRACKET",(90, 85, "LeftBracket") },
            { "RIGHT_BRACKET",(91, 86, "RightBracket") },
            { "LEFT_BRACE",(92, 87, "LeftBrace") },
            { "RIGHT_BRACE",(93, 88, "RightBrace") },
            { "BRACKET_PAIR",(94, 185, "BracketPair") },
            { "BRACE_PAIR",(95, 186, "BracePair") },
            { "STRAIGHT_CONNECTOR_1",(96, 32, "StraightConnector1") },
            { "BENT_CONNECTOR_2",(97, 33, "BentConnector2") },
            { "BENT_CONNECTOR_3",(98, 34, "BentConnector3") },
            { "BENT_CONNECTOR_4",(99, 35, "BentConnector4") },
            { "BENT_CONNECTOR_5",(100, 36, "BentConnector5") },
            { "CURVED_CONNECTOR_2",(101, 37, "CurvedConnector2") },
            { "CURVED_CONNECTOR_3",(102, 38, "CurvedConnector3") },
            { "CURVED_CONNECTOR_4",(103, 39, "CurvedConnector4") },
            { "CURVED_CONNECTOR_5",(104, 40, "CurvedConnector5") },
            { "CALLOUT_1",(105, 41, "Callout1") },
            { "CALLOUT_2",(106, 42, "Callout2") },
            { "CALLOUT_3",(107, 43, "Callout3") },
            { "ACCENT_CALLOUT_1",(108, 44, "AccentCallout1") },
            { "ACCENT_CALLOUT_2",(109, 45, "AccentCallout2") },
            { "ACCENT_CALLOUT_3",(110, 46, "AccentCallout3") },
            { "BORDER_CALLOUT_1",(111, 47, "BorderCallout1") },
            { "BORDER_CALLOUT_2",(112, 48, "BorderCallout2") },
            { "BORDER_CALLOUT_3",(113, 49, "BorderCallout3") },
            { "ACCENT_BORDER_CALLOUT_1",(114, 50, "AccentBorderCallout1") },
            { "ACCENT_BORDER_CALLOUT_2",(115, 51, "AccentBorderCallout2") },
            { "ACCENT_BORDER_CALLOUT_3",(116, 52, "AccentBorderCallout3") },
            { "WEDGE_RECT_CALLOUT",(117, 61, "WedgeRectCallout") },
            { "WEDGE_ROUND_RECT_CALLOUT",(118, 62, "WedgeRRectCallout") },
            { "WEDGE_ELLIPSE_CALLOUT",(119, 63, "WedgeEllipseCallout") },
            { "CLOUD_CALLOUT",(120, 106, "CloudCallout") },
            { "CLOUD",(121, -1, null) },
            { "RIBBON",(122, 53, "Ribbon") },
            { "RIBBON_2",(123, 54, "Ribbon2") },
            { "ELLIPSE_RIBBON",(124, 107, "EllipseRibbon") },
            { "ELLIPSE_RIBBON_2",(125, 108, "EllipseRibbon2") },
            { "LEFT_RIGHT_RIBBON",(126, -1, null) },
            { "VERTICAL_SCROLL",(127, 97, "VerticalScroll") },
            { "HORIZONTAL_SCROLL",(128, 98, "HorizontalScroll") },
            { "WAVE",(129, 64, "Wave") },
            { "DOUBLE_WAVE",(130, 188, "DoubleWave") },
            { "PLUS",(131, 11, "Plus") },
            { "FLOW_CHART_PROCESS",(132, 109, "FlowChartProcess") },
            { "FLOW_CHART_DECISION",(133, 110, "FlowChartDecision") },
            { "FLOW_CHART_INPUT_OUTPUT",(134, 111, "FlowChartInputOutput") },
            { "FLOW_CHART_PREDEFINED_PROCESS",(135, 112, "FlowChartPredefinedProcess") },
            { "FLOW_CHART_INTERNAL_STORAGE",(136, 113, "FlowChartInternalStorage") },
            { "FLOW_CHART_DOCUMENT",(137, 114, "FlowChartDocument") },
            { "FLOW_CHART_MULTIDOCUMENT",(138, 115, "FlowChartMultidocument") },
            { "FLOW_CHART_TERMINATOR",(139, 116, "FlowChartTerminator") },
            { "FLOW_CHART_PREPARATION",(140, 117, "FlowChartPreparation") },
            { "FLOW_CHART_MANUAL_INPUT",(141, 118, "FlowChartManualInput") },
            { "FLOW_CHART_MANUAL_OPERATION",(142, 119, "FlowChartManualOperation") },
            { "FLOW_CHART_CONNECTOR",(143, 120, "FlowChartConnector") },
            { "FLOW_CHART_PUNCHED_CARD",(144, 121, "FlowChartPunchedCard") },
            { "FLOW_CHART_PUNCHED_TAPE",(145, 122, "FlowChartPunchedTape") },
            { "FLOW_CHART_SUMMING_JUNCTION",(146, 123, "FlowChartSummingJunction") },
            { "FLOW_CHART_OR",(147, 124, "FlowChartOr") },
            { "FLOW_CHART_COLLATE",(148, 125, "FlowChartCollate") },
            { "FLOW_CHART_SORT",(149, 126, "FlowChartSort") },
            { "FLOW_CHART_EXTRACT",(150, 127, "FlowChartExtract") },
            { "FLOW_CHART_MERGE",(151, 128, "FlowChartMerge") },
            { "FLOW_CHART_OFFLINE_STORAGE",(152, 129, "FlowChartOfflineStorage") },
            { "FLOW_CHART_ONLINE_STORAGE",(153, 130, "FlowChartOnlineStorage") },
            { "FLOW_CHART_MAGNETIC_TAPE",(154, 131, "FlowChartMagneticTape") },
            { "FLOW_CHART_MAGNETIC_DISK",(155, 132, "FlowChartMagneticDisk") },
            { "FLOW_CHART_MAGNETIC_DRUM",(156, 133, "FlowChartMagneticDrum") },
            { "FLOW_CHART_DISPLAY",(157, 134, "FlowChartDisplay") },
            { "FLOW_CHART_DELAY",(158, 135, "FlowChartDelay") },
            { "FLOW_CHART_ALTERNATE_PROCESS",(159, 176, "FlowChartAlternateProcess") },
            { "FLOW_CHART_OFFPAGE_CONNECTOR",(160, 177, "FlowChartOffpageConnector") },
            { "ACTION_BUTTON_BLANK",(161, 189, "ActionButtonBlank") },
            { "ACTION_BUTTON_HOME",(162, 190, "ActionButtonHome") },
            { "ACTION_BUTTON_HELP",(163, 191, "ActionButtonHelp") },
            { "ACTION_BUTTON_INFORMATION",(164, 192, "ActionButtonInformation") },
            { "ACTION_BUTTON_FORWARD_NEXT",(165, 193, "ActionButtonForwardNext") },
            { "ACTION_BUTTON_BACK_PREVIOUS",(166, 194, "ActionButtonBackPrevious") },
            { "ACTION_BUTTON_END",(167, 195, "ActionButtonEnd") },
            { "ACTION_BUTTON_BEGINNING",(168, 196, "ActionButtonBeginning") },
            { "ACTION_BUTTON_RETURN",(169, 197, "ActionButtonReturn") },
            { "ACTION_BUTTON_DOCUMENT",(170, 198, "ActionButtonDocument") },
            { "ACTION_BUTTON_SOUND",(171, 199, "ActionButtonSound") },
            { "ACTION_BUTTON_MOVIE",(172, 200, "ActionButtonMovie") },
            { "GEAR_6",(173, -1, null) },
            { "GEAR_9",(174, -1, null) },
            { "FUNNEL",(175, -1, null) },
            { "MATH_PLUS",(176, -1, null) },
     MATH_MINUS(177, -1, null),
     MATH_MULTIPLY(178, -1, null),
     MATH_DIVIDE(179, -1, null),
     MATH_EQUAL(180, -1, null),
     MATH_NOT_EQUAL(181, -1, null),
     CORNER_TABS(182, -1, null),
     SQUARE_TABS(183, -1, null),
     PLAQUE_TABS(184, -1, null),
     CHART_X(185, -1, null),
     CHART_STAR(186, -1, null),
     CHART_PLUS(187, -1, null),
	    // below are shape types only found in native
     NOTCHED_CIRCULAR_ARROW(-1, 100, "NotchedCircularArrow"),
     THICK_ARROW(-1, 14, "ThickArrow"),
     BALLOON(-1, 17, "Balloon"),
     TEXT_SIMPLE(-1, 24, "TextSimple"),
     TEXT_OCTAGON(-1, 25, "TextOctagon"),
     TEXT_HEXAGON(-1, 26, "TextHexagon"),
     TEXT_CURVE(-1, 27, "TextCurve"),
     TEXT_WAVE(-1, 28, "TextWave"),
     TEXT_RING(-1, 29, "TextRing"),
     TEXT_ON_CURVE(-1, 30, "TextOnCurve"),
     TEXT_ON_RING(-1, 31, "TextOnRing"),
     TEXT_PLAIN_TEXT(-1, 136, "TextPlainText"),
     TEXT_STOP(-1, 137, "TextStop"),
     TEXT_TRIANGLE(-1, 138, "TextTriangle"),
     TEXT_TRIANGLE_INVERTED(-1, 139, "TextTriangleInverted"),
     TEXT_CHEVRON(-1, 140, "TextChevron"),
     TEXT_CHEVRON_INVERTED(-1, 141, "TextChevronInverted"),
     TEXT_RING_INSIDE(-1, 142, "TextRingInside"),
     TEXT_RING_OUTSIDE(-1, 143, "TextRingOutside"),
     TEXT_ARCH_UP_CURVE(-1, 144, "TextArchUpCurve"),
     TEXT_ARCH_DOWN_CURVE(-1, 145, "TextArchDownCurve"),
     TEXT_CIRCLE_CURVE(-1, 146, "TextCircleCurve"),
     TEXT_BUTTON_CURVE(-1, 147, "TextButtonCurve"),
     TEXT_ARCH_UP_POUR(-1, 148, "TextArchUpPour"),
     TEXT_ARCH_DOWN_POUR(-1, 149, "TextArchDownPour"),
     TEXT_CIRCLE_POUR(-1, 150, "TextCirclePour"),
     TEXT_BUTTON_POUR(-1, 151, "TextButtonPour"),
     TEXT_CURVE_UP(-1, 152, "TextCurveUp"),
     TEXT_CURVE_DOWN(-1, 153, "TextCurveDown"),
     TEXT_CASCADE_UP(-1, 154, "TextCascadeUp"),
     TEXT_CASCADE_DOWN(-1, 155, "TextCascadeDown"),
     TEXT_WAVE_1(-1, 156, "TextWave1"),
     TEXT_WAVE_2(-1, 157, "TextWave2"),
     TEXT_WAVE_3(-1, 158, "TextWave3"),
     TEXT_WAVE_4(-1, 159, "TextWave4"),
     TEXT_INFLATE(-1, 160, "TextInflate"),
     TEXT_DEFLATE(-1, 161, "TextDeflate"),
     TEXT_INFLATE_BOTTOM(-1, 162, "TextInflateBottom"),
     TEXT_DEFLATE_BOTTOM(-1, 163, "TextDeflateBottom"),
     TEXT_INFLATE_TOP(-1, 164, "TextInflateTop"),
     TEXT_DEFLATE_TOP(-1, 165, "TextDeflateTop"),
     TEXT_DEFLATE_INFLATE(-1, 166, "TextDeflateInflate"),
     TEXT_DEFLATE_INFLATE_DEFLATE(-1, 167, "TextDeflateInflateDeflate"),
     TEXT_FADE_RIGHT(-1, 168, "TextFadeRight"),
     TEXT_FADE_LEFT(-1, 169, "TextFadeLeft"),
     TEXT_FADE_UP(-1, 170, "TextFadeUp"),
     TEXT_FADE_DOWN(-1, 171, "TextFadeDown"),
     TEXT_SLANT_UP(-1, 172, "TextSlantUp"),
     TEXT_SLANT_DOWN(-1, 173, "TextSlantDown"),
     TEXT_CAN_UP(-1, 174, "TextCanUp"),
     TEXT_CAN_DOWN(-1, 175, "TextCanDown"),
     CALLOUT_90(-1, 178, "Callout90"),
     ACCENT_CALLOUT_90(-1, 179, "AccentCallout90"),
     BORDER_CALLOUT_90(-1, 180, "BorderCallout90"),
     ACCENT_BORDER_CALLOUT_90(-1, 181, "AccentBorderCallout90"),
     HOST_CONTROL(-1, 201, "HostControl"),
     TEXT_BOX(-1, 202, "TextBox")
		};


		public ShapeType(int ooxmlId, int nativeId, string nativeName)
		{
			this.ooxmlId = ooxmlId;
			this.nativeId = nativeId;
			this.nativeName = nativeName;
		}
	
	    /** name of the presetShapeDefinit(i)on entry */
	    public String getOoxmlName()
		{
			if (this == SEAL) return STAR_16.getOoxmlName();
			if (ooxmlId == -1)
			{
				return (name().startsWith("TEXT")) ? RECT.getOoxmlName() : null;
			}

	         StringBuilder sb = new StringBuilder();
			bool toLower = true;
			for (char ch : name().toCharArray())
			{
				if (ch == '_')
				{
					toLower = false;
					continue;
				}
				sb.append(toLower ? StringUtil.toLowerCase(ch) : StringUtil.toUpperCase(ch));
				toLower = true;
			}

         return sb.toString();
		}
	    
	    public static ShapeType forId(int id, bool isOoxmlId)
		{
			// exemption for #60294
			if (!isOoxmlId && id == 0x0FFF)
			{
				return NOT_PRIMITIVE;
			}
			
	         for (ShapeType t : values())
			{
				if ((isOoxmlId && t.ooxmlId == id) ||
					(!isOoxmlId && t.nativeId == id)) return t;
			}
			throw new IllegalArgumentException("Unknown shape type: " + id);
		}
	}
}