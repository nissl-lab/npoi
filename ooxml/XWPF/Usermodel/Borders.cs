/* ====================================================================
   Licensed to the Apache Software Foundation =(ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   =(the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.XWPF.UserModel
{
    using System;




    /**
     * Specifies all types of borders which can be specified for WordProcessingML
     * objects which have a border. Borders can be Separated into two types:
     * <ul>
     * <li> Line borders: which specify a pattern to be used when Drawing a line around the
     * specified object.
     * </li>
     * <li> Art borders: which specify a repeated image to be used
     * when Drawing a border around the specified object. Line borders may be
     * specified on any object which allows a border, however, art borders may only
     * be used as a border at the page level - the borders under the pgBorders
     * element
     *</li>
     * </ul>
     * @author Gisella Bronzetti
     */
    public enum Borders
    {

        NIL = (1),

        NONE = (2),

        /**
         * Specifies a line border consisting of a single line around the parent
         * object.
         */
        SINGLE = (3),

        THICK = (4),

        DOUBLE = (5),

        DOTTED = (6),

        DASHED = (7),

        DOT_DASH = (8),

        DOT_DOT_DASH = (9),

        TRIPLE = (10),

        THIN_THICK_SMALL_GAP = (11),

        THICK_THIN_SMALL_GAP = (12),

        THIN_THICK_THIN_SMALL_GAP = (13),

        THIN_THICK_MEDIUM_GAP = (14),

        THICK_THIN_MEDIUM_GAP = (15),

        THIN_THICK_THIN_MEDIUM_GAP = (16),

        THIN_THICK_LARGE_GAP = (17),

        THICK_THIN_LARGE_GAP = (18),

        THIN_THICK_THIN_LARGE_GAP = (19),

        WAVE = (20),

        DOUBLE_WAVE = (21),

        DASH_SMALL_GAP = (22),

        DASH_DOT_STROKED = (23),

        THREE_D_EMBOSS = (24),

        THREE_D_ENGRAVE = (25),

        OUTSET = (26),

        INSET = (27),

        /**
         * Specifies an art border consisting of a repeated image of an apple
         */
        APPLES = (28),

        /**
         * Specifies an art border consisting of a repeated image of a shell pattern
         */
        ARCHED_SCALLOPS = (29),

        /**
         * Specifies an art border consisting of a repeated image of a baby pacifier
         */
        BABY_PACIFIER = (30),

        /**
         * Specifies an art border consisting of a repeated image of a baby rattle
         */
        BABY_RATTLE = (31),

        /**
         * Specifies an art border consisting of a repeated image of a Set of
         * balloons
         */
        BALLOONS_3_COLORS = (32),

        /**
         * Specifies an art border consisting of a repeated image of a hot air
         * balloon
         */
        BALLOONS_HOT_AIR = (33),

        /**
         * Specifies an art border consisting of a repeating image of a black and
         * white background.
         */
        BASIC_BLACK_DASHES = (34),

        /**
         * Specifies an art border consisting of a repeating image of a black dot on
         * a white background.
         */
        BASIC_BLACK_DOTS = (35),

        /**
         * Specifies an art border consisting of a repeating image of a black and
         * white background
         */
        BASIC_BLACK_SQUARES = (36),

        /**
         * Specifies an art border consisting of a repeating image of a black and
         * white background.
         */
        BASIC_THIN_LINES = (37),

        /**
         * Specifies an art border consisting of a repeating image of a black and
         * white background.
         */
        BASIC_WHITE_DASHES = (38),

        /**
         * Specifies an art border consisting of a repeating image of a white dot on
         * a black background.
         */
        BASIC_WHITE_DOTS = (39),

        /**
         * Specifies an art border consisting of a repeating image of a black and
         * white background.
         */
        BASIC_WHITE_SQUARES = (40),

        /**
         * Specifies an art border consisting of a repeating image of a black and
         * white background.
         */
        BASIC_WIDE_INLINE = (41),

        /**
         * Specifies an art border consisting of a repeating image of a black and
         * white background
         */
        BASIC_WIDE_MIDLINE = (42),

        /**
         * Specifies an art border consisting of a repeating image of a black and
         * white background
         */
        BASIC_WIDE_OUTLINE = (43),

        /**
         * Specifies an art border consisting of a repeated image of bats
         */
        BATS = (44),

        /**
         * Specifies an art border consisting of repeating images of birds
         */
        BIRDS = (45),

        /**
         * Specifies an art border consisting of a repeated image of birds flying
         */
        BIRDS_FLIGHT = (46),

        /**
         * Specifies an art border consisting of a repeated image of a cabin
         */
        CABINS = (47),

        /**
         * Specifies an art border consisting of a repeated image of a piece of cake
         */
        CAKE_SLICE = (48),

        /**
         * Specifies an art border consisting of a repeated image of candy corn
         */
        CANDY_CORN = (49),

        /**
         * Specifies an art border consisting of a repeated image of a knot work
         * pattern
         */
        CELTIC_KNOTWORK = (50),

        /**
         * Specifies an art border consisting of a banner.
         * <p>
         * If the border is on the left or right, no border is displayed.
         * </p>
         */
        CERTIFICATE_BANNER = (51),

        /**
         * Specifies an art border consisting of a repeating image of a chain link
         * pattern.
         */
        CHAIN_LINK = (52),

        /**
         * Specifies an art border consisting of a repeated image of a champagne
         * bottle
         */
        CHAMPAGNE_BOTTLE = (53),

        /**
         * Specifies an art border consisting of repeating images of a compass
         */
        CHECKED_BAR_BLACK = (54),

        /**
         * Specifies an art border consisting of a repeating image of a colored
         * pattern.
         */
        CHECKED_BAR_COLOR = (55),

        /**
         * Specifies an art border consisting of a repeated image of a Checkerboard
         */
        CHECKERED = (56),

        /**
         * Specifies an art border consisting of a repeated image of a Christmas
         * tree
         */
        CHRISTMAS_TREE = (57),

        /**
         * Specifies an art border consisting of repeating images of lines and
         * circles
         */
        CIRCLES_LINES = (58),

        /**
         * Specifies an art border consisting of a repeated image of a rectangular
         * pattern
         */
        CIRCLES_RECTANGLES = (59),

        /**
         * Specifies an art border consisting of a repeated image of a wave
         */
        CLASSICAL_WAVE = (60),

        /**
         * Specifies an art border consisting of a repeated image of a clock
         */
        CLOCKS = (61),

        /**
         * Specifies an art border consisting of repeating images of a compass
         */
        COMPASS = (62),

        /**
         * Specifies an art border consisting of a repeated image of confetti
         */
        CONFETTI = (63),

        /**
         * Specifies an art border consisting of a repeated image of confetti
         */
        CONFETTI_GRAYS = (64),

        /**
         * Specifies an art border consisting of a repeated image of confetti
         */
        CONFETTI_OUTLINE = (65),

        /**
         * Specifies an art border consisting of a repeated image of confetti
         * streamers
         */
        CONFETTI_STREAMERS = (66),

        /**
         * Specifies an art border consisting of a repeated image of confetti
         */
        CONFETTI_WHITE = (67),

        /**
         * Specifies an art border consisting of a repeated image
         */
        CORNER_TRIANGLES = (68),

        /**
         * Specifies an art border consisting of a dashed line
         */
        COUPON_CUTOUT_DASHES = (69),

        /**
         * Specifies an art border consisting of a dotted line
         */
        COUPON_CUTOUT_DOTS = (70),

        /**
         * Specifies an art border consisting of a repeated image of a maze-like
         * pattern
         */
        CRAZY_MAZE = (71),

        /**
         * Specifies an art border consisting of a repeated image of a butterfly
         */
        CREATURES_BUTTERFLY = (72),

        /**
         * Specifies an art border consisting of a repeated image of a fish
         */
        CREATURES_FISH = (73),

        /**
         * Specifies an art border consisting of repeating images of insects.
         */
        CREATURES_INSECTS = (74),

        /**
         * Specifies an art border consisting of a repeated image of a ladybug
         */
        CREATURES_LADY_BUG = (75),

        /**
         * Specifies an art border consisting of repeating images of a cross-stitch
         * pattern
         */
        CROSS_STITCH = (76),

        /**
         * Specifies an art border consisting of a repeated image of Cupid
         */
        CUP = (77),

        DECO_ARCH = (78),

        DECO_ARCH_COLOR = (79),

        DECO_BLOCKS = (80),

        DIAMONDS_GRAY = (81),

        DOUBLE_D = (82),

        DOUBLE_DIAMONDS = (83),

        EARTH_1 = (84),

        EARTH_2 = (85),

        ECLIPSING_SQUARES_1 = (86),

        ECLIPSING_SQUARES_2 = (87),

        EGGS_BLACK = (88),

        FANS = (89),

        FILM = (90),

        FIRECRACKERS = (91),

        FLOWERS_BLOCK_PRINT = (92),

        FLOWERS_DAISIES = (93),

        FLOWERS_MODERN_1 = (94),

        FLOWERS_MODERN_2 = (95),

        FLOWERS_PANSY = (96),

        FLOWERS_RED_ROSE = (97),

        FLOWERS_ROSES = (98),

        FLOWERS_TEACUP = (99),

        FLOWERS_TINY = (100),

        GEMS = (101),

        GINGERBREAD_MAN = (102),

        GRADIENT = (103),

        HANDMADE_1 = (104),

        HANDMADE_2 = (105),

        HEART_BALLOON = (106),

        HEART_GRAY = (107),

        HEARTS = (108),

        HEEBIE_JEEBIES = (109),

        HOLLY = (110),

        HOUSE_FUNKY = (111),

        HYPNOTIC = (112),

        ICE_CREAM_CONES = (113),

        LIGHT_BULB = (114),

        LIGHTNING_1 = (115),

        LIGHTNING_2 = (116),

        MAP_PINS = (117),

        MAPLE_LEAF = (118),

        MAPLE_MUFFINS = (119),

        MARQUEE = (120),

        MARQUEE_TOOTHED = (121),

        MOONS = (122),

        MOSAIC = (123),

        MUSIC_NOTES = (124),

        NORTHWEST = (125),

        OVALS = (126),

        PACKAGES = (127),

        PALMS_BLACK = (128),

        PALMS_COLOR = (129),

        PAPER_CLIPS = (130),

        PAPYRUS = (131),

        PARTY_FAVOR = (132),

        PARTY_GLASS = (133),

        PENCILS = (134),

        PEOPLE = (135),

        PEOPLE_WAVING = (136),

        PEOPLE_HATS = (137),

        POINSETTIAS = (138),

        POSTAGE_STAMP = (139),

        PUMPKIN_1 = (140),

        PUSH_PIN_NOTE_2 = (141),

        PUSH_PIN_NOTE_1 = (142),

        PYRAMIDS = (143),

        PYRAMIDS_ABOVE = (144),

        QUADRANTS = (145),

        RINGS = (146),

        SAFARI = (147),

        SAWTOOTH = (148),

        SAWTOOTH_GRAY = (149),

        SCARED_CAT = (150),

        SEATTLE = (151),

        SHADOWED_SQUARES = (152),

        SHARKS_TEETH = (153),

        SHOREBIRD_TRACKS = (154),

        SKYROCKET = (155),

        SNOWFLAKE_FANCY = (156),

        SNOWFLAKES = (157),

        SOMBRERO = (158),

        SOUTHWEST = (159),

        STARS = (160),

        STARS_TOP = (161),

        STARS_3_D = (162),

        STARS_BLACK = (163),

        STARS_SHADOWED = (164),

        SUN = (165),

        SWIRLIGIG = (166),

        TORN_PAPER = (167),

        TORN_PAPER_BLACK = (168),

        TREES = (169),

        TRIANGLE_PARTY = (170),

        TRIANGLES = (171),

        TRIBAL_1 = (172),

        TRIBAL_2 = (173),

        TRIBAL_3 = (174),

        TRIBAL_4 = (175),

        TRIBAL_5 = (176),

        TRIBAL_6 = (177),

        TWISTED_LINES_1 = (178),

        TWISTED_LINES_2 = (179),

        VINE = (180),

        WAVELINE = (181),

        WEAVING_ANGLES = (182),

        WEAVING_BRAID = (183),

        WEAVING_RIBBON = (184),

        WEAVING_STRIPS = (185),

        WHITE_FLOWERS = (186),

        WOODWORK = (187),

        X_ILLUSIONS = (188),

        ZANY_TRIANGLES = (189),

        ZIG_ZAG = (190),

        ZIG_ZAG_STITCH = (191)

        //private int value;

        //private Borders(int val) {
        //    value = val;
        //}

        //public int GetValue() {
        //    return value;
        //}

        //private static Dictionary<int, Borders> imap = new Dictionary<int, Borders>();
        //static {
        //    foreach =(Borders p in values=()) {
        //        imap.Put=(Int32.ValueOf=(p.Value), p);
        //    }
        //}

        //public static Borders ValueOf(int type) {
        //    Borders pBorder = imap.Get(Int32.ValueOf=(type));
        //    if (pBorder == null) {
        //        throw new ArgumentException("Unknown paragraph border: " + type);
        //    }
        //    return pBorder;
        //}
    }
}