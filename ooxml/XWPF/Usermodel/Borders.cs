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

        Nil = (1),

        None = (2),

        /**
         * Specifies a line border consisting of a single line around the parent
         * object.
         */
        Single = (3),

        Thick = (4),

        Double = (5),

        Dotted = (6),

        Dashed = (7),

        DotDash = (8),

        DotDotDash = (9),

        Triple = (10),

        ThinThickSmallGap = (11),

        ThickThinSmallGap = (12),

        ThinThickThinSmallGap = (13),

        ThinThickMediumGap = (14),

        ThickThinMediumGap = (15),

        ThinThickThinMediumGap = (16),

        ThinThickLargeGap = (17),

        ThickThinLargeGap = (18),

        ThinThickThinLargeGap = (19),

        Wave = (20),

        DoubleWave = (21),

        DashSmallGap = (22),

        DashDotStroked = (23),

        ThreeDEmboss = (24),

        ThreeDEngrave = (25),

        Outset = (26),

        Inset = (27),

        /**
         * specifies an art border consisting of a repeated image of an apple
         */
        Apples = (28),

        /**
         * specifies an art border consisting of a repeated image of a shell pattern
         */
        ArchedScallops = (29),

        /**
         * specifies an art border consisting of a repeated image of a baby pacifier
         */
        BabyPacifier = (30),

        /**
         * specifies an art border consisting of a repeated image of a baby rattle
         */
        BabyRattle = (31),

        /**
         * specifies an art border consisting of a repeated image of a set of
         * balloons
         */
        Balloons3Colors = (32),

        /**
         * specifies an art border consisting of a repeated image of a hot air
         * balloon
         */
        BalloonsHotAir = (33),

        /**
         * specifies an art border consisting of a repeating image of a black and
         * white background.
         */
        BasicBlackDashes = (34),

        /**
         * specifies an art border consisting of a repeating image of a black dot on
         * a white background.
         */
        BasicBlackDots = (35),

        /**
         * specifies an art border consisting of a repeating image of a black and
         * white background
         */
        BasicBlackSquares = (36),

        /**
         * specifies an art border consisting of a repeating image of a black and
         * white background.
         */
        BasicThinLines = (37),

        /**
         * specifies an art border consisting of a repeating image of a black and
         * white background.
         */
        BasicWhiteDashes = (38),

        /**
         * specifies an art border consisting of a repeating image of a white dot on
         * a black background.
         */
        BasicWhiteDots = (39),

        /**
         * specifies an art border consisting of a repeating image of a black and
         * white background.
         */
        BasicWhiteSquares = (40),

        /**
         * specifies an art border consisting of a repeating image of a black and
         * white background.
         */
        BasicWideInline = (41),

        /**
         * specifies an art border consisting of a repeating image of a black and
         * white background
         */
        BasicWideMidline = (42),

        /**
         * specifies an art border consisting of a repeating image of a black and
         * white background
         */
        BasicWideOutline = (43),

        /**
         * specifies an art border consisting of a repeated image of bats
         */
        Bats = (44),

        /**
         * specifies an art border consisting of repeating images of birds
         */
        Birds = (45),

        /**
         * specifies an art border consisting of a repeated image of birds flying
         */
        BirdsFlight = (46),

        /**
         * specifies an art border consisting of a repeated image of a cabin
         */
        Cabins = (47),

        /**
         * specifies an art border consisting of a repeated image of a piece of cake
         */
        CakeSlice = (48),

        /**
         * specifies an art border consisting of a repeated image of candy corn
         */
        CandyCorn = (49),

        /**
         * specifies an art border consisting of a repeated image of a knot work
         * pattern
         */
        CelticKnotwork = (50),

        /**
         * specifies an art border consisting of a banner.
         * <p>
         * if the border is on the left or right, no border is displayed.
         * </p>
         */
        CertificateBanner = (51),

        /**
         * specifies an art border consisting of a repeating image of a chain link
         * pattern.
         */
        ChainLink = (52),

        /**
         * specifies an art border consisting of a repeated image of a champagne
         * bottle
         */
        ChampagneBottle = (53),

        /**
         * specifies an art border consisting of repeating images of a compass
         */
        CheckedBarBlack = (54),

        /**
         * specifies an art border consisting of a repeating image of a colored
         * pattern.
         */
        CheckedBarColor = (55),

        /**
         * specifies an art border consisting of a repeated image of a checkerboard
         */
        Checkered = (56),

        /**
         * specifies an art border consisting of a repeated image of a christmas
         * tree
         */
        ChristmasTree = (57),

        /**
         * specifies an art border consisting of repeating images of lines and
         * circles
         */
        CirclesLines = (58),

        /**
         * specifies an art border consisting of a repeated image of a rectangular
         * pattern
         */
        CirclesRectangles = (59),

        /**
         * specifies an art border consisting of a repeated image of a wave
         */
        ClassicalWave = (60),

        /**
         * specifies an art border consisting of a repeated image of a clock
         */
        Clocks = (61),

        /**
         * specifies an art border consisting of repeating images of a compass
         */
        Compass = (62),

        /**
         * specifies an art border consisting of a repeated image of confetti
         */
        Confetti = (63),

        /**
         * specifies an art border consisting of a repeated image of confetti
         */
        ConfettiGrays = (64),

        /**
         * specifies an art border consisting of a repeated image of confetti
         */
        ConfettiOutline = (65),

        /**
         * specifies an art border consisting of a repeated image of confetti
         * streamers
         */
        ConfettiStreamers = (66),

        /**
         * specifies an art border consisting of a repeated image of confetti
         */
        ConfettiWhite = (67),

        /**
         * specifies an art border consisting of a repeated image
         */
        CornerTriangles = (68),

        /**
         * specifies an art border consisting of a dashed line
         */
        CouponCutoutDashes = (69),

        /**
         * specifies an art border consisting of a dotted line
         */
        CouponCutoutDots = (70),

        /**
         * specifies an art border consisting of a repeated image of a maze-like
         * pattern
         */
        CrazyMaze = (71),

        /**
         * specifies an art border consisting of a repeated image of a butterfly
         */
        CreaturesButterfly = (72),

        /**
         * specifies an art border consisting of a repeated image of a fish
         */
        CreaturesFish = (73),

        /**
         * specifies an art border consisting of repeating images of insects.
         */
        CreaturesInsects = (74),

        /**
         * specifies an art border consisting of a repeated image of a ladybug
         */
        CreaturesLadyBug = (75),

        /**
         * specifies an art border consisting of repeating images of a cross-stitch
         * pattern
         */
        CrossStitch = (76),

        /**
         * specifies an art border consisting of a repeated image of cupid
         */
        Cup = (77),

        DecoArch = (78),

        DecoArchColor = (79),

        DecoBlocks = (80),

        DiamondsGray = (81),

        DoubleD = (82),

        DoubleDiamonds = (83),

        Earth1 = (84),

        Earth2 = (85),

        EclipsingSquares1 = (86),

        EclipsingSquares2 = (87),

        EggsBlack = (88),

        Fans = (89),

        Film = (90),

        Firecrackers = (91),

        FlowersBlockPrint = (92),

        FlowersDaisies = (93),

        FlowersModern1 = (94),

        FlowersModern2 = (95),

        FlowersPansy = (96),

        FlowersRedRose = (97),

        FlowersRoses = (98),

        FlowersTeacup = (99),

        FlowersTiny = (100),

        Gems = (101),

        GingerbreadMan = (102),

        Gradient = (103),

        Handmade1 = (104),

        Handmade2 = (105),

        HeartBalloon = (106),

        HeartGray = (107),

        Hearts = (108),

        HeebieJeebies = (109),

        Holly = (110),

        HouseFunky = (111),

        Hypnotic = (112),

        IceCreamCones = (113),

        LightBulb = (114),

        Lightning1 = (115),

        Lightning2 = (116),

        MapPins = (117),

        MapleLeaf = (118),

        MapleMuffins = (119),

        Marquee = (120),

        MarqueeToothed = (121),

        Moons = (122),

        Mosaic = (123),

        MusicNotes = (124),

        Northwest = (125),

        Ovals = (126),

        Packages = (127),

        PalmsBlack = (128),

        PalmsColor = (129),

        PaperClips = (130),

        Papyrus = (131),

        PartyFavor = (132),

        PartyGlass = (133),

        Pencils = (134),

        People = (135),

        PeopleWaving = (136),

        PeopleHats = (137),

        Poinsettias = (138),

        PostageStamp = (139),

        Pumpkin1 = (140),

        PushPinNote2 = (141),

        PushPinNote1 = (142),

        Pyramids = (143),

        PyramidsAbove = (144),

        Quadrants = (145),

        Rings = (146),

        Safari = (147),

        Sawtooth = (148),

        SawtoothGray = (149),

        ScaredCat = (150),

        Seattle = (151),

        ShadowedSquares = (152),

        SharksTeeth = (153),

        ShorebirdTracks = (154),

        Skyrocket = (155),

        SnowflakeFancy = (156),

        Snowflakes = (157),

        Sombrero = (158),

        Southwest = (159),

        Stars = (160),

        StarsTop = (161),

        Stars3D = (162),

        StarsBlack = (163),

        StarsShadowed = (164),

        Sun = (165),

        Swirligig = (166),

        TornPaper = (167),

        TornPaperBlack = (168),

        Trees = (169),

        TriangleParty = (170),

        Triangles = (171),

        Tribal1 = (172),

        Tribal2 = (173),

        Tribal3 = (174),

        Tribal4 = (175),

        Tribal5 = (176),

        Tribal6 = (177),

        TwistedLines1 = (178),

        TwistedLines2 = (179),

        Vine = (180),

        Waveline = (181),

        WeavingAngles = (182),

        WeavingBraid = (183),

        WeavingRibbon = (184),

        WeavingStrips = (185),

        WhiteFlowers = (186),

        Woodwork = (187),

        XIllusions = (188),

        ZanyTriangles = (189),

        ZigZag = (190),

        ZigZagStitch = (191)


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