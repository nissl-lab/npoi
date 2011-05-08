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
namespace NPOI.SS.UserModel
{

    /**
     * All known types of automatic shapes in DrawingML
     *
     * @author Yegor Kozlov
     */
    public class ShapeTypes
    {
        public static int LINE = 1;
        public static int LINE_INV = 2;
        public static int TRIANGLE = 3;
        public static int RT_TRIANGLE = 4;
        public static int RECT = 5;
        public static int DIAMOND = 6;
        public static int PARALLELOGRAM = 7;
        public static int TRAPEZOID = 8;
        public static int NON_ISOSCELES_TRAPEZOID = 9;
        public static int PENTAGON = 10;
        public static int HEXAGON = 11;
        public static int HEPTAGON = 12;
        public static int OCTAGON = 13;
        public static int DECAGON = 14;
        public static int DODECAGON = 15;
        public static int STAR_4 = 16;
        public static int STAR_5 = 17;
        public static int STAR_6 = 18;
        public static int STAR_7 = 19;
        public static int STAR_8 = 20;
        public static int STAR_10 = 21;
        public static int STAR_12 = 22;
        public static int STAR_16 = 23;
        public static int STAR_24 = 24;
        public static int STAR_32 = 25;
        public static int ROUND_RECT = 26;
        public static int ROUND_1_RECT = 27;
        public static int ROUND_2_SAME_RECT = 28;
        public static int ROUND_2_DIAG_RECT = 29;
        public static int SNIP_ROUND_RECT = 30;
        public static int SNIP_1_RECT = 31;
        public static int SNIP_2_SAME_RECT = 32;
        public static int SNIP_2_DIAG_RECT = 33;
        public static int PLAQUE = 34;
        public static int ELLIPSE = 35;
        public static int TEARDROP = 36;
        public static int HOME_PLATE = 37;
        public static int CHEVRON = 38;
        public static int PIE_WEDGE = 39;
        public static int PIE = 40;
        public static int BLOCK_ARC = 41;
        public static int DONUT = 42;
        public static int NO_SMOKING = 43;
        public static int RIGHT_ARROW = 44;
        public static int LEFT_ARROW = 45;
        public static int UP_ARROW = 46;
        public static int DOWN_ARROW = 47;
        public static int STRIPED_RIGHT_ARROW = 48;
        public static int NOTCHED_RIGHT_ARROW = 49;
        public static int BENT_UP_ARROW = 50;
        public static int LEFT_RIGHT_ARROW = 51;
        public static int UP_DOWN_ARROW = 52;
        public static int LEFT_UP_ARROW = 53;
        public static int LEFT_RIGHT_UP_ARROW = 54;
        public static int QUAD_ARROW = 55;
        public static int LEFT_ARROW_CALLOUT = 56;
        public static int RIGHT_ARROW_CALLOUT = 57;
        public static int UP_ARROW_CALLOUT = 58;
        public static int DOWN_ARROW_CALLOUT = 59;
        public static int LEFT_RIGHT_ARROW_CALLOUT = 60;
        public static int UP_DOWN_ARROW_CALLOUT = 61;
        public static int QUAD_ARROW_CALLOUT = 62;
        public static int BENT_ARROW = 63;
        public static int UTURN_ARROW = 64;
        public static int CIRCULAR_ARROW = 65;
        public static int LEFT_CIRCULAR_ARROW = 66;
        public static int LEFT_RIGHT_CIRCULAR_ARROW = 67;
        public static int CURVED_RIGHT_ARROW = 68;
        public static int CURVED_LEFT_ARROW = 69;
        public static int CURVED_UP_ARROW = 70;
        public static int CURVED_DOWN_ARROW = 71;
        public static int SWOOSH_ARROW = 72;
        public static int CUBE = 73;
        public static int CAN = 74;
        public static int LIGHTNING_BOLT = 75;
        public static int HEART = 76;
        public static int SUN = 77;
        public static int MOON = 78;
        public static int SMILEY_FACE = 79;
        public static int IRREGULAR_SEAL_1 = 80;
        public static int IRREGULAR_SEAL_2 = 81;
        public static int FOLDED_CORNER = 82;
        public static int BEVEL = 83;
        public static int FRAME = 84;
        public static int HALF_FRAME = 85;
        public static int CORNER = 86;
        public static int DIAG_STRIPE = 87;
        public static int CHORD = 88;
        public static int ARC = 89;
        public static int LEFT_BRACKET = 90;
        public static int RIGHT_BRACKET = 91;
        public static int LEFT_BRACE = 92;
        public static int RIGHT_BRACE = 93;
        public static int BRACKET_PAIR = 94;
        public static int BRACE_PAIR = 95;
        public static int STRAIGHT_CONNECTOR_1 = 96;
        public static int BENT_CONNECTOR_2 = 97;
        public static int BENT_CONNECTOR_3 = 98;
        public static int BENT_CONNECTOR_4 = 99;
        public static int BENT_CONNECTOR_5 = 100;
        public static int CURVED_CONNECTOR_2 = 101;
        public static int CURVED_CONNECTOR_3 = 102;
        public static int CURVED_CONNECTOR_4 = 103;
        public static int CURVED_CONNECTOR_5 = 104;
        public static int CALLOUT_1 = 105;
        public static int CALLOUT_2 = 106;
        public static int CALLOUT_3 = 107;
        public static int ACCENT_CALLOUT_1 = 108;
        public static int ACCENT_CALLOUT_2 = 109;
        public static int ACCENT_CALLOUT_3 = 110;
        public static int BORDER_CALLOUT_1 = 111;
        public static int BORDER_CALLOUT_2 = 112;
        public static int BORDER_CALLOUT_3 = 113;
        public static int ACCENT_BORDER_CALLOUT_1 = 114;
        public static int ACCENT_BORDER_CALLOUT_2 = 115;
        public static int ACCENT_BORDER_CALLOUT_3 = 116;
        public static int WEDGE_RECT_CALLOUT = 117;
        public static int WEDGE_ROUND_RECT_CALLOUT = 118;
        public static int WEDGE_ELLIPSE_CALLOUT = 119;
        public static int CLOUD_CALLOUT = 120;
        public static int CLOUD = 121;
        public static int RIBBON = 122;
        public static int RIBBON_2 = 123;
        public static int ELLIPSE_RIBBON = 124;
        public static int ELLIPSE_RIBBON_2 = 125;
        public static int LEFT_RIGHT_RIBBON = 126;
        public static int VERTICAL_SCROLL = 127;
        public static int HORIZONTAL_SCROLL = 128;
        public static int WAVE = 129;
        public static int DOUBLE_WAVE = 130;
        public static int PLUS = 131;
        public static int FLOW_CHART_PROCESS = 132;
        public static int FLOW_CHART_DECISION = 133;
        public static int FLOW_CHART_INPUT_OUTPUT = 134;
        public static int FLOW_CHART_PREDEFINED_PROCESS = 135;
        public static int FLOW_CHART_INTERNAL_STORAGE = 136;
        public static int FLOW_CHART_DOCUMENT = 137;
        public static int FLOW_CHART_MULTIDOCUMENT = 138;
        public static int FLOW_CHART_TERMINATOR = 139;
        public static int FLOW_CHART_PREPARATION = 140;
        public static int FLOW_CHART_MANUAL_INPUT = 141;
        public static int FLOW_CHART_MANUAL_OPERATION = 142;
        public static int FLOW_CHART_CONNECTOR = 143;
        public static int FLOW_CHART_PUNCHED_CARD = 144;
        public static int FLOW_CHART_PUNCHED_TAPE = 145;
        public static int FLOW_CHART_SUMMING_JUNCTION = 146;
        public static int FLOW_CHART_OR = 147;
        public static int FLOW_CHART_COLLATE = 148;
        public static int FLOW_CHART_SORT = 149;
        public static int FLOW_CHART_EXTRACT = 150;
        public static int FLOW_CHART_MERGE = 151;
        public static int FLOW_CHART_OFFLINE_STORAGE = 152;
        public static int FLOW_CHART_ONLINE_STORAGE = 153;
        public static int FLOW_CHART_MAGNETIC_TAPE = 154;
        public static int FLOW_CHART_MAGNETIC_DISK = 155;
        public static int FLOW_CHART_MAGNETIC_DRUM = 156;
        public static int FLOW_CHART_DISPLAY = 157;
        public static int FLOW_CHART_DELAY = 158;
        public static int FLOW_CHART_ALTERNATE_PROCESS = 159;
        public static int FLOW_CHART_OFFPAGE_CONNECTOR = 160;
        public static int ACTION_BUTTON_BLANK = 161;
        public static int ACTION_BUTTON_HOME = 162;
        public static int ACTION_BUTTON_HELP = 163;
        public static int ACTION_BUTTON_INFORMATION = 164;
        public static int ACTION_BUTTON_FORWARD_NEXT = 165;
        public static int ACTION_BUTTON_BACK_PREVIOUS = 166;
        public static int ACTION_BUTTON_END = 167;
        public static int ACTION_BUTTON_BEGINNING = 168;
        public static int ACTION_BUTTON_RETURN = 169;
        public static int ACTION_BUTTON_DOCUMENT = 170;
        public static int ACTION_BUTTON_SOUND = 171;
        public static int ACTION_BUTTON_MOVIE = 172;
        public static int GEAR_6 = 173;
        public static int GEAR_9 = 174;
        public static int FUNNEL = 175;
        public static int MATH_PLUS = 176;
        public static int MATH_MINUS = 177;
        public static int MATH_MULTIPLY = 178;
        public static int MATH_DIVIDE = 179;
        public static int MATH_EQUAL = 180;
        public static int MATH_NOT_EQUAL = 181;
        public static int CORNER_TABS = 182;
        public static int SQUARE_TABS = 183;
        public static int PLAQUE_TABS = 184;
        public static int CHART_X = 185;
        public static int CHART_STAR = 186;
        public static int CHART_PLUS = 187;
    }

}