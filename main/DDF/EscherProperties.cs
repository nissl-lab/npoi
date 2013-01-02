
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

namespace NPOI.DDF
{
    using System;
    using System.Collections;

    /// <summary>
    /// Provides a list of all known escher properties including the description and
    /// type.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class EscherProperties
    {

        #region Property constants
        public const short TRANSFORM__ROTATION = 4;
        public const short PROTECTION__LOCKROTATION = 119;
        public const short PROTECTION__LOCKASPECTRATIO = 120;
        public const short PROTECTION__LOCKPOSITION = 121;
        public const short PROTECTION__LOCKAGAINSTSELECT = 122;
        public const short PROTECTION__LOCKCROPPING = 123;
        public const short PROTECTION__LOCKVERTICES = 124;
        public const short PROTECTION__LOCKTEXT = 125;
        public const short PROTECTION__LOCKADJUSTHANDLES = 126;
        public const short PROTECTION__LOCKAGAINSTGROUPING = 127;
        public const short TEXT__TEXTID = 128;
        public const short TEXT__TEXTLEFT = 129;
        public const short TEXT__TEXTTOP = 130;
        public const short TEXT__TEXTRIGHT = 131;
        public const short TEXT__TEXTBOTTOM = 132;
        public const short TEXT__WRAPTEXT = 133;
        public const short TEXT__SCALETEXT = 134;
        public const short TEXT__ANCHORTEXT = 135;
        public const short TEXT__TEXTFLOW = 136;
        public const short TEXT__FONTROTATION = 137;
        public const short TEXT__IDOFNEXTSHAPE = 138;
        public const short TEXT__BIDIR = 139;
        public const short TEXT__SINGLECLICKSELECTS = 187;
        public const short TEXT__USEHOSTMARGINS = 188;
        public const short TEXT__ROTATETEXTWITHSHAPE = 189;
        public const short TEXT__SIZESHAPETOFITTEXT = 190;
        public const short TEXT__SIZE_TEXT_TO_FIT_SHAPE = 191;
        public const short GEOTEXT__UNICODE = 192;
        public const short GEOTEXT__RTFTEXT = 193;
        public const short GEOTEXT__ALIGNMENTONCURVE = 194;
        public const short GEOTEXT__DEFAULTPOINTSIZE = 195;
        public const short GEOTEXT__TEXTSPACING = 196;
        public const short GEOTEXT__FONTFAMILYNAME = 197;
        public const short GEOTEXT__REVERSEROWORDER = 240;
        public const short GEOTEXT__HASTEXTEFFECT = 241;
        public const short GEOTEXT__ROTATECHARACTERS = 242;
        public const short GEOTEXT__KERNCHARACTERS = 243;
        public const short GEOTEXT__TIGHTORTRACK = 244;
        public const short GEOTEXT__STRETCHTOFITSHAPE = 245;
        public const short GEOTEXT__CHARBOUNDINGBOX = 246;
        public const short GEOTEXT__SCALETEXTONPATH = 247;
        public const short GEOTEXT__STRETCHCHARHEIGHT = 248;
        public const short GEOTEXT__NOMEASUREALONGPATH = 249;
        public const short GEOTEXT__BOLDFONT = 250;
        public const short GEOTEXT__ITALICFONT = 251;
        public const short GEOTEXT__UNDERLINEFONT = 252;
        public const short GEOTEXT__SHADOWFONT = 253;
        public const short GEOTEXT__SMALLCAPSFONT = 254;
        public const short GEOTEXT__STRIKETHROUGHFONT = 255;
        public const short BLIP__CROPFROMTOP = 256;
        public const short BLIP__CROPFROMBOTTOM = 257;
        public const short BLIP__CROPFROMLEFT = 258;
        public const short BLIP__CROPFROMRIGHT = 259;
        public const short BLIP__BLIPTODISPLAY = 260;
        public const short BLIP__BLIPFILENAME = 261;
        public const short BLIP__BLIPFLAGS = 262;
        public const short BLIP__TRANSPARENTCOLOR = 263;
        public const short BLIP__CONTRASTSetTING = 264;
        public const short BLIP__BRIGHTNESSSetTING = 265;
        public const short BLIP__GAMMA = 266;
        public const short BLIP__PICTUREID = 267;
        public const short BLIP__DOUBLEMOD = 268;
        public const short BLIP__PICTUREFillMOD = 269;
        public const short BLIP__PICTURELINE = 270;
        public const short BLIP__PRINTBLIP = 271;
        public const short BLIP__PRINTBLIPFILENAME = 272;
        public const short BLIP__PRINTFLAGS = 273;
        public const short BLIP__NOHITTESTPICTURE = 316;
        public const short BLIP__PICTUREGRAY = 317;
        public const short BLIP__PICTUREBILEVEL = 318;
        public const short BLIP__PICTUREACTIVE = 319;
        public const short GEOMETRY__LEFT = 320;
        public const short GEOMETRY__TOP = 321;
        public const short GEOMETRY__RIGHT = 322;
        public const short GEOMETRY__BOTTOM = 323;
        public const short GEOMETRY__SHAPEPATH = 324;
        public const short GEOMETRY__VERTICES = 325;
        public const short GEOMETRY__SEGMENTINFO = 326;
        public const short GEOMETRY__ADJUSTVALUE = 327;
        public const short GEOMETRY__ADJUST2VALUE = 328;
        public const short GEOMETRY__ADJUST3VALUE = 329;
        public const short GEOMETRY__ADJUST4VALUE = 330;
        public const short GEOMETRY__ADJUST5VALUE = 331;
        public const short GEOMETRY__ADJUST6VALUE = 332;
        public const short GEOMETRY__ADJUST7VALUE = 333;
        public const short GEOMETRY__ADJUST8VALUE = 334;
        public const short GEOMETRY__ADJUST9VALUE = 335;
        public const short GEOMETRY__ADJUST10VALUE = 336;
        public const short GEOMETRY__SHADOWok = 378;
        public const short GEOMETRY__3DOK = 379;
        public const short GEOMETRY__LINEOK = 380;
        public const short GEOMETRY__GEOTEXTOK = 381;
        public const short GEOMETRY__FILLSHADESHAPEOK = 382;
        public const short GEOMETRY__FILLOK = 383;
        public const short FILL__FILLTYPE = 384;
        public const short FILL__FILLCOLOR = 385;
        public const short FILL__FILLOPACITY = 386;
        public const short FILL__FILLBACKCOLOR = 387;
        public const short FILL__BACKOPACITY = 388;
        public const short FILL__CRMOD = 389;
        public const short FILL__PATTERNTEXTURE = 390;
        public const short FILL__BLIPFILENAME = 391;
        public const short FILL__BLIPFLAGS = 392;
        public const short FILL__WIDTH = 393;
        public const short FILL__HEIGHT = 394;
        public const short FILL__ANGLE = 395;
        public const short FILL__FOCUS = 396;
        public const short FILL__TOLEFT = 397;
        public const short FILL__TOTOP = 398;
        public const short FILL__TORIGHT = 399;
        public const short FILL__TOBOTTOM = 400;
        public const short FILL__RECTLEFT = 401;
        public const short FILL__RECTTOP = 402;
        public const short FILL__RECTRIGHT = 403;
        public const short FILL__RECTBOTTOM = 404;
        public const short FILL__DZTYPE = 405;
        public const short FILL__SHADEPRESet = 406;
        public const short FILL__SHADECOLORS = 407;
        public const short FILL__ORIGINX = 408;
        public const short FILL__ORIGINY = 409;
        public const short FILL__SHAPEORIGINX = 410;
        public const short FILL__SHAPEORIGINY = 411;
        public const short FILL__SHADETYPE = 412;
        public const short FILL__FILLED = 443;
        public const short FILL__HITTESTFILL = 444;
        public const short FILL__SHAPE = 445;
        public const short FILL__USERECT = 446;
        public const short FILL__NOFILLHITTEST = 447;
        public const short LINESTYLE__COLOR = 448;
        public const short LINESTYLE__OPACITY = 449;
        public const short LINESTYLE__BACKCOLOR = 450;
        public const short LINESTYLE__CRMOD = 451;
        public const short LINESTYLE__LINETYPE = 452;
        public const short LINESTYLE__FILLBLIP = 453;
        public const short LINESTYLE__FILLBLIPNAME = 454;
        public const short LINESTYLE__FILLBLIPFLAGS = 455;
        public const short LINESTYLE__FILLWIDTH = 456;
        public const short LINESTYLE__FILLHEIGHT = 457;
        public const short LINESTYLE__FILLDZTYPE = 458;
        public const short LINESTYLE__LINEWIDTH = 459;
        public const short LINESTYLE__LINEMITERLIMIT = 460;
        public const short LINESTYLE__LINESTYLE = 461;
        public const short LINESTYLE__LINEDASHING = 462;
        public const short LINESTYLE__LINEDASHSTYLE = 463;
        public const short LINESTYLE__LINESTARTARROWHEAD = 464;
        public const short LINESTYLE__LINEENDARROWHEAD = 465;
        public const short LINESTYLE__LINESTARTARROWWIDTH = 466;
        public const short LINESTYLE__LINEESTARTARROWLength = 467;
        public const short LINESTYLE__LINEENDARROWWIDTH = 468;
        public const short LINESTYLE__LINEENDARROWLength = 469;
        public const short LINESTYLE__LINEJOINSTYLE = 470;
        public const short LINESTYLE__LINEENDCAPSTYLE = 471;
        public const short LINESTYLE__ARROWHEADSOK = 507;
        public const short LINESTYLE__ANYLINE = 508;
        public const short LINESTYLE__HITLINETEST = 509;
        public const short LINESTYLE__LINEFILLSHAPE = 510;
        public const short LINESTYLE__NOLINEDRAWDASH = 511;
        public const short SHADOWSTYLE__TYPE = 512;
        public const short SHADOWSTYLE__COLOR = 513;
        public const short SHADOWSTYLE__HIGHLIGHT = 514;
        public const short SHADOWSTYLE__CRMOD = 515;
        public const short SHADOWSTYLE__OPACITY = 516;
        public const short SHADOWSTYLE__OFFSetX = 517;
        public const short SHADOWSTYLE__OFFSetY = 518;
        public const short SHADOWSTYLE__SECONDOFFSetX = 519;
        public const short SHADOWSTYLE__SECONDOFFSetY = 520;
        public const short SHADOWSTYLE__SCALEXTOX = 521;
        public const short SHADOWSTYLE__SCALEYTOX = 522;
        public const short SHADOWSTYLE__SCALEXTOY = 523;
        public const short SHADOWSTYLE__SCALEYTOY = 524;
        public const short SHADOWSTYLE__PERSPECTIVEX = 525;
        public const short SHADOWSTYLE__PERSPECTIVEY = 526;
        public const short SHADOWSTYLE__WEIGHT = 527;
        public const short SHADOWSTYLE__ORIGINX = 528;
        public const short SHADOWSTYLE__ORIGINY = 529;
        public const short SHADOWSTYLE__SHADOW = 574;
        public const short SHADOWSTYLE__SHADOWOBSURED = 575;
        public const short PERSPECTIVE__TYPE = 576;
        public const short PERSPECTIVE__OFFSetX = 577;
        public const short PERSPECTIVE__OFFSetY = 578;
        public const short PERSPECTIVE__SCALEXTOX = 579;
        public const short PERSPECTIVE__SCALEYTOX = 580;
        public const short PERSPECTIVE__SCALEXTOY = 581;
        public const short PERSPECTIVE__SCALEYTOY = 582;
        public const short PERSPECTIVE__PERSPECTIVEX = 583;
        public const short PERSPECTIVE__PERSPECTIVEY = 584;
        public const short PERSPECTIVE__WEIGHT = 585;
        public const short PERSPECTIVE__ORIGINX = 586;
        public const short PERSPECTIVE__ORIGINY = 587;
        public const short PERSPECTIVE__PERSPECTIVEON = 639;
        public const short THREED__SPECULARAMOUNT = 640;
        public const short THREED__DIFFUSEAMOUNT = 661;
        public const short THREED__SHININESS = 662;
        public const short THREED__EDGetHICKNESS = 663;
        public const short THREED__EXTRUDEFORWARD = 664;
        public const short THREED__EXTRUDEBACKWARD = 665;
        public const short THREED__EXTRUDEPLANE = 666;
        public const short THREED__EXTRUSIONCOLOR = 667;
        public const short THREED__CRMOD = 648;
        public const short THREED__3DEFFECT = 700;
        public const short THREED__METALLIC = 701;
        public const short THREED__USEEXTRUSIONCOLOR = 702;
        public const short THREED__LIGHTFACE = 703;
        public const short THREEDSTYLE__YROTATIONANGLE = 704;
        public const short THREEDSTYLE__XROTATIONANGLE = 705;
        public const short THREEDSTYLE__ROTATIONAXISX = 706;
        public const short THREEDSTYLE__ROTATIONAXISY = 707;
        public const short THREEDSTYLE__ROTATIONAXISZ = 708;
        public const short THREEDSTYLE__ROTATIONANGLE = 709;
        public const short THREEDSTYLE__ROTATIONCENTERX = 710;
        public const short THREEDSTYLE__ROTATIONCENTERY = 711;
        public const short THREEDSTYLE__ROTATIONCENTERZ = 712;
        public const short THREEDSTYLE__RENDERMODE = 713;
        public const short THREEDSTYLE__TOLERANCE = 714;
        public const short THREEDSTYLE__XVIEWPOINT = 715;
        public const short THREEDSTYLE__YVIEWPOINT = 716;
        public const short THREEDSTYLE__ZVIEWPOINT = 717;
        public const short THREEDSTYLE__ORIGINX = 718;
        public const short THREEDSTYLE__ORIGINY = 719;
        public const short THREEDSTYLE__SKEWANGLE = 720;
        public const short THREEDSTYLE__SKEWAMOUNT = 721;
        public const short THREEDSTYLE__AMBIENTINTENSITY = 722;
        public const short THREEDSTYLE__KEYX = 723;
        public const short THREEDSTYLE__KEYY = 724;
        public const short THREEDSTYLE__KEYZ = 725;
        public const short THREEDSTYLE__KEYINTENSITY = 726;
        public const short THREEDSTYLE__FillX = 727;
        public const short THREEDSTYLE__FillY = 728;
        public const short THREEDSTYLE__FillZ = 729;
        public const short THREEDSTYLE__FillINTENSITY = 730;
        public const short THREEDSTYLE__CONSTRAINROTATION = 763;
        public const short THREEDSTYLE__ROTATIONCENTERAUTO = 764;
        public const short THREEDSTYLE__PARALLEL = 765;
        public const short THREEDSTYLE__KEYHARSH = 766;
        public const short THREEDSTYLE__FillHARSH = 767;
        public const short SHAPE__MASTER = 769;
        public const short SHAPE__CONNECTORSTYLE = 771;
        public const short SHAPE__BLACKANDWHITESetTINGS = 772;
        public const short SHAPE__WMODEPUREBW = 773;
        public const short SHAPE__WMODEBW = 774;
        public const short SHAPE__OLEICON = 826;
        public const short SHAPE__PREFERRELATIVERESIZE = 827;
        public const short SHAPE__LOCKSHAPETYPE = 828;
        public const short SHAPE__DELETEATTACHEDOBJECT = 830;
        public const short SHAPE__BACKGROUNDSHAPE = 831;
        public const short CALLOUT__CALLOUTTYPE = 832;
        public const short CALLOUT__XYCALLOUTGAP = 833;
        public const short CALLOUT__CALLOUTANGLE = 834;
        public const short CALLOUT__CALLOUTDROPTYPE = 835;
        public const short CALLOUT__CALLOUTDROPSPECIFIED = 836;
        public const short CALLOUT__CALLOUTLENGTHSPECIFIED = 837;
        public const short CALLOUT__ISCALLOUT = 889;
        public const short CALLOUT__CALLOUTACCENTBAR = 890;
        public const short CALLOUT__CALLOUTTEXTBORDER = 891;
        public const short CALLOUT__CALLOUTMINUSX = 892;
        public const short CALLOUT__CALLOUTMINUSY = 893;
        public const short CALLOUT__DROPAUTO = 894;
        public const short CALLOUT__LENGTHSPECIFIED = 895;

	    public const short GROUPSHAPE__SHAPENAME = 0x0380;
	    public const short GROUPSHAPE__DESCRIPTION = 0x0381;
        public const short GROUPSHAPE__HYPERLINK = 0x0382;
	    public const short GROUPSHAPE__WRAPPOLYGONVERTICES = 0x0383;
	    public const short GROUPSHAPE__WRAPDISTLEFT = 0x0384;
	    public const short GROUPSHAPE__WRAPDISTTOP = 0x0385;
	    public const short GROUPSHAPE__WRAPDISTRIGHT = 0x0386;
	    public const short GROUPSHAPE__WRAPDISTBOTTOM = 0x0387;
        public const short GROUPSHAPE__REGROUPID = 0x0388;
        public const short GROUPSHAPE__UNUSED906 = 0x038A;
        public const short GROUPSHAPE__TOOLTIP = 0x038D;
        public const short GROUPSHAPE__SCRIPT = 0x038E;
        public const short GROUPSHAPE__POSH = 0x038F;
        public const short GROUPSHAPE__POSRELH = 0x0390;
        public const short GROUPSHAPE__POSV = 0x0391;
        public const short GROUPSHAPE__POSRELV = 0x0392;
        public const short GROUPSHAPE__HR_PCT = 0x0393;
        public const short GROUPSHAPE__HR_ALIGN = 0x0394;
        public const short GROUPSHAPE__HR_HEIGHT = 0x0395;
        public const short GROUPSHAPE__HR_WIDTH = 0x0396;
        public const short GROUPSHAPE__SCRIPTEXT = 0x0397;
        public const short GROUPSHAPE__SCRIPTLANG = 0x0398;
        public const short GROUPSHAPE__BORDERTOPCOLOR = 0x039B;
        public const short GROUPSHAPE__BORDERLEFTCOLOR = 0x039C;
        public const short GROUPSHAPE__BORDERBOTTOMCOLOR = 0x039D;
        public const short GROUPSHAPE__BORDERRIGHTCOLOR = 0x039E;
        public const short GROUPSHAPE__TABLEPROPERTIES = 0x039F;
        public const short GROUPSHAPE__TABLEROWPROPERTIES = 0x03A0;
        public const short GROUPSHAPE__WEBBOT = 0x03A5;
        public const short GROUPSHAPE__METROBLOB = 0x03A9;
        public const short GROUPSHAPE__ZORDER = 0x03AA;
        public const short GROUPSHAPE__FLAGS = 0x03BF;
	    public const short GROUPSHAPE__EDITEDWRAP = 953;
	    public const short GROUPSHAPE__BEHINDDOCUMENT = 954;
	    public const short GROUPSHAPE__ONDBLCLICKNOTIFY = 955;
	    public const short GROUPSHAPE__ISBUTTON = 956;
	    public const short GROUPSHAPE__1DADJUSTMENT = 957;
	    public const short GROUPSHAPE__HIDDEN = 958;
	    public const short GROUPSHAPE__PRINT = 959;
        #endregion

        private static Hashtable properties;

        /// <summary>
        /// Inits the props.
        /// </summary>
        private static void InitProps()
        {
            if (properties == null)
            {
                properties = new Hashtable();
                AddProp(TRANSFORM__ROTATION, GetData("transform.rotation"));
                AddProp(PROTECTION__LOCKROTATION, GetData("protection.lockrotation"));
                AddProp(PROTECTION__LOCKASPECTRATIO, GetData("protection.lockaspectratio"));
                AddProp(PROTECTION__LOCKPOSITION, GetData("protection.lockposition"));
                AddProp(PROTECTION__LOCKAGAINSTSELECT, GetData("protection.lockagainstselect"));
                AddProp(PROTECTION__LOCKCROPPING, GetData("protection.lockcropping"));
                AddProp(PROTECTION__LOCKVERTICES, GetData("protection.lockvertices"));
                AddProp(PROTECTION__LOCKTEXT, GetData("protection.locktext"));
                AddProp(PROTECTION__LOCKADJUSTHANDLES, GetData("protection.lockadjusthandles"));
                AddProp(PROTECTION__LOCKAGAINSTGROUPING, GetData("protection.lockagainstgrouping", EscherPropertyMetaData.TYPE_bool));
                AddProp(TEXT__TEXTID, GetData("text.textid"));
                AddProp(TEXT__TEXTLEFT, GetData("text.textleft"));
                AddProp(TEXT__TEXTTOP, GetData("text.texttop"));
                AddProp(TEXT__TEXTRIGHT, GetData("text.textright"));
                AddProp(TEXT__TEXTBOTTOM, GetData("text.textbottom"));
                AddProp(TEXT__WRAPTEXT, GetData("text.wraptext"));
                AddProp(TEXT__SCALETEXT, GetData("text.scaletext"));
                AddProp(TEXT__ANCHORTEXT, GetData("text.anchortext"));
                AddProp(TEXT__TEXTFLOW, GetData("text.textflow"));
                AddProp(TEXT__FONTROTATION, GetData("text.fontrotation"));
                AddProp(TEXT__IDOFNEXTSHAPE, GetData("text.idofnextshape"));
                AddProp(TEXT__BIDIR, GetData("text.bidir"));
                AddProp(TEXT__SINGLECLICKSELECTS, GetData("text.singleclickselects"));
                AddProp(TEXT__USEHOSTMARGINS, GetData("text.usehostmargins"));
                AddProp(TEXT__ROTATETEXTWITHSHAPE, GetData("text.rotatetextwithshape"));
                AddProp(TEXT__SIZESHAPETOFITTEXT, GetData("text.sizeshapetofittext"));
                AddProp(TEXT__SIZE_TEXT_TO_FIT_SHAPE, GetData("text.sizetexttofitshape", EscherPropertyMetaData.TYPE_bool));
                AddProp(GEOTEXT__UNICODE, GetData("geotext.unicode"));
                AddProp(GEOTEXT__RTFTEXT, GetData("geotext.rtftext"));
                AddProp(GEOTEXT__ALIGNMENTONCURVE, GetData("geotext.alignmentoncurve"));
                AddProp(GEOTEXT__DEFAULTPOINTSIZE, GetData("geotext.defaultpointsize"));
                AddProp(GEOTEXT__TEXTSPACING, GetData("geotext.textspacing"));
                AddProp(GEOTEXT__FONTFAMILYNAME, GetData("geotext.fontfamilyname"));
                AddProp(GEOTEXT__REVERSEROWORDER, GetData("geotext.reverseroworder"));
                AddProp(GEOTEXT__HASTEXTEFFECT, GetData("geotext.hastexteffect"));
                AddProp(GEOTEXT__ROTATECHARACTERS, GetData("geotext.rotatecharacters"));
                AddProp(GEOTEXT__KERNCHARACTERS, GetData("geotext.kerncharacters"));
                AddProp(GEOTEXT__TIGHTORTRACK, GetData("geotext.tightortrack"));
                AddProp(GEOTEXT__STRETCHTOFITSHAPE, GetData("geotext.stretchtofitshape"));
                AddProp(GEOTEXT__CHARBOUNDINGBOX, GetData("geotext.charboundingbox"));
                AddProp(GEOTEXT__SCALETEXTONPATH, GetData("geotext.scaletextonpath"));
                AddProp(GEOTEXT__STRETCHCHARHEIGHT, GetData("geotext.stretchcharheight"));
                AddProp(GEOTEXT__NOMEASUREALONGPATH, GetData("geotext.nomeasurealongpath"));
                AddProp(GEOTEXT__BOLDFONT, GetData("geotext.boldfont"));
                AddProp(GEOTEXT__ITALICFONT, GetData("geotext.italicfont"));
                AddProp(GEOTEXT__UNDERLINEFONT, GetData("geotext.underlinefont"));
                AddProp(GEOTEXT__SHADOWFONT, GetData("geotext.shadowfont"));
                AddProp(GEOTEXT__SMALLCAPSFONT, GetData("geotext.smallcapsfont"));
                AddProp(GEOTEXT__STRIKETHROUGHFONT, GetData("geotext.strikethroughfont"));
                AddProp(BLIP__CROPFROMTOP, GetData("blip.cropfromtop"));
                AddProp(BLIP__CROPFROMBOTTOM, GetData("blip.cropfrombottom"));
                AddProp(BLIP__CROPFROMLEFT, GetData("blip.cropfromleft"));
                AddProp(BLIP__CROPFROMRIGHT, GetData("blip.cropfromright"));
                AddProp(BLIP__BLIPTODISPLAY, GetData("blip.bliptodisplay"));
                AddProp(BLIP__BLIPFILENAME, GetData("blip.blipfilename"));
                AddProp(BLIP__BLIPFLAGS, GetData("blip.blipflags"));
                AddProp(BLIP__TRANSPARENTCOLOR, GetData("blip.transparentcolor"));
                AddProp(BLIP__CONTRASTSetTING, GetData("blip.contrastSetting"));
                AddProp(BLIP__BRIGHTNESSSetTING, GetData("blip.brightnessSetting"));
                AddProp(BLIP__GAMMA, GetData("blip.gamma"));
                AddProp(BLIP__PICTUREID, GetData("blip.pictureid"));
                AddProp(BLIP__DOUBLEMOD, GetData("blip.doublemod"));
                AddProp(BLIP__PICTUREFillMOD, GetData("blip.pictureFillmod"));
                AddProp(BLIP__PICTURELINE, GetData("blip.pictureline"));
                AddProp(BLIP__PRINTBLIP, GetData("blip.printblip"));
                AddProp(BLIP__PRINTBLIPFILENAME, GetData("blip.printblipfilename"));
                AddProp(BLIP__PRINTFLAGS, GetData("blip.printflags"));
                AddProp(BLIP__NOHITTESTPICTURE, GetData("blip.nohittestpicture"));
                AddProp(BLIP__PICTUREGRAY, GetData("blip.picturegray"));
                AddProp(BLIP__PICTUREBILEVEL, GetData("blip.picturebilevel"));
                AddProp(BLIP__PICTUREACTIVE, GetData("blip.pictureactive"));
                AddProp(GEOMETRY__LEFT, GetData("geometry.left"));
                AddProp(GEOMETRY__TOP, GetData("geometry.top"));
                AddProp(GEOMETRY__RIGHT, GetData("geometry.right"));
                AddProp(GEOMETRY__BOTTOM, GetData("geometry.bottom"));
                AddProp(GEOMETRY__SHAPEPATH, GetData("geometry.shapepath", EscherPropertyMetaData.TYPE_SHAPEPATH));
                AddProp(GEOMETRY__VERTICES, GetData("geometry.vertices", EscherPropertyMetaData.TYPE_ARRAY));
                AddProp(GEOMETRY__SEGMENTINFO, GetData("geometry.segmentinfo", EscherPropertyMetaData.TYPE_ARRAY));
                AddProp(GEOMETRY__ADJUSTVALUE, GetData("geometry.adjustvalue"));
                AddProp(GEOMETRY__ADJUST2VALUE, GetData("geometry.adjust2value"));
                AddProp(GEOMETRY__ADJUST3VALUE, GetData("geometry.adjust3value"));
                AddProp(GEOMETRY__ADJUST4VALUE, GetData("geometry.adjust4value"));
                AddProp(GEOMETRY__ADJUST5VALUE, GetData("geometry.adjust5value"));
                AddProp(GEOMETRY__ADJUST6VALUE, GetData("geometry.adjust6value"));
                AddProp(GEOMETRY__ADJUST7VALUE, GetData("geometry.adjust7value"));
                AddProp(GEOMETRY__ADJUST8VALUE, GetData("geometry.adjust8value"));
                AddProp(GEOMETRY__ADJUST9VALUE, GetData("geometry.adjust9value"));
                AddProp(GEOMETRY__ADJUST10VALUE, GetData("geometry.adjust10value"));
                AddProp(GEOMETRY__SHADOWok, GetData("geometry.shadowOK"));
                AddProp(GEOMETRY__3DOK, GetData("geometry.3dok"));
                AddProp(GEOMETRY__LINEOK, GetData("geometry.lineok"));
                AddProp(GEOMETRY__GEOTEXTOK, GetData("geometry.geotextok"));
                AddProp(GEOMETRY__FILLSHADESHAPEOK, GetData("geometry.fillshadeshapeok"));
                AddProp(GEOMETRY__FILLOK, GetData("geometry.fillok", EscherPropertyMetaData.TYPE_bool));
                AddProp(FILL__FILLTYPE, GetData("fill.filltype"));
                AddProp(FILL__FILLCOLOR, GetData("fill.fillcolor", EscherPropertyMetaData.TYPE_RGB));
                AddProp(FILL__FILLOPACITY, GetData("fill.fillopacity"));
                AddProp(FILL__FILLBACKCOLOR, GetData("fill.fillbackcolor", EscherPropertyMetaData.TYPE_RGB));
                AddProp(FILL__BACKOPACITY, GetData("fill.backopacity"));
                AddProp(FILL__CRMOD, GetData("fill.crmod"));
                AddProp(FILL__PATTERNTEXTURE, GetData("fill.patterntexture"));
                AddProp(FILL__BLIPFILENAME, GetData("fill.blipfilename"));
                AddProp(FILL__BLIPFLAGS, GetData("fill.blipflags"));
                AddProp(FILL__WIDTH, GetData("fill.width"));
                AddProp(FILL__HEIGHT, GetData("fill.height"));
                AddProp(FILL__ANGLE, GetData("fill.angle"));
                AddProp(FILL__FOCUS, GetData("fill.focus"));
                AddProp(FILL__TOLEFT, GetData("fill.toleft"));
                AddProp(FILL__TOTOP, GetData("fill.totop"));
                AddProp(FILL__TORIGHT, GetData("fill.toright"));
                AddProp(FILL__TOBOTTOM, GetData("fill.tobottom"));
                AddProp(FILL__RECTLEFT, GetData("fill.rectleft"));
                AddProp(FILL__RECTTOP, GetData("fill.recttop"));
                AddProp(FILL__RECTRIGHT, GetData("fill.rectright"));
                AddProp(FILL__RECTBOTTOM, GetData("fill.rectbottom"));
                AddProp(FILL__DZTYPE, GetData("fill.dztype"));
                AddProp(FILL__SHADEPRESet, GetData("fill.shadepReset"));
                AddProp(FILL__SHADECOLORS, GetData("fill.shadecolors", EscherPropertyMetaData.TYPE_ARRAY));
                AddProp(FILL__ORIGINX, GetData("fill.originx"));
                AddProp(FILL__ORIGINY, GetData("fill.originy"));
                AddProp(FILL__SHAPEORIGINX, GetData("fill.shapeoriginx"));
                AddProp(FILL__SHAPEORIGINY, GetData("fill.shapeoriginy"));
                AddProp(FILL__SHADETYPE, GetData("fill.shadetype"));
                AddProp(FILL__FILLED, GetData("fill.filled"));
                AddProp(FILL__HITTESTFILL, GetData("fill.hittestfill"));
                AddProp(FILL__SHAPE, GetData("fill.shape"));
                AddProp(FILL__USERECT, GetData("fill.userect"));
                AddProp(FILL__NOFILLHITTEST, GetData("fill.nofillhittest", EscherPropertyMetaData.TYPE_bool));
                AddProp(LINESTYLE__COLOR, GetData("linestyle.color", EscherPropertyMetaData.TYPE_RGB));
                AddProp(LINESTYLE__OPACITY, GetData("linestyle.opacity"));
                AddProp(LINESTYLE__BACKCOLOR, GetData("linestyle.backcolor", EscherPropertyMetaData.TYPE_RGB));
                AddProp(LINESTYLE__CRMOD, GetData("linestyle.crmod"));
                AddProp(LINESTYLE__LINETYPE, GetData("linestyle.linetype"));
                AddProp(LINESTYLE__FILLBLIP, GetData("linestyle.fillblip"));
                AddProp(LINESTYLE__FILLBLIPNAME, GetData("linestyle.fillblipname"));
                AddProp(LINESTYLE__FILLBLIPFLAGS, GetData("linestyle.fillblipflags"));
                AddProp(LINESTYLE__FILLWIDTH, GetData("linestyle.fillwidth"));
                AddProp(LINESTYLE__FILLHEIGHT, GetData("linestyle.fillheight"));
                AddProp(LINESTYLE__FILLDZTYPE, GetData("linestyle.filldztype"));
                AddProp(LINESTYLE__LINEWIDTH, GetData("linestyle.linewidth"));
                AddProp(LINESTYLE__LINEMITERLIMIT, GetData("linestyle.linemiterlimit"));
                AddProp(LINESTYLE__LINESTYLE, GetData("linestyle.linestyle"));
                AddProp(LINESTYLE__LINEDASHING, GetData("linestyle.linedashing"));
                AddProp(LINESTYLE__LINEDASHSTYLE, GetData("linestyle.linedashstyle", EscherPropertyMetaData.TYPE_ARRAY));
                AddProp(LINESTYLE__LINESTARTARROWHEAD, GetData("linestyle.linestartarrowhead"));
                AddProp(LINESTYLE__LINEENDARROWHEAD, GetData("linestyle.lineendarrowhead"));
                AddProp(LINESTYLE__LINESTARTARROWWIDTH, GetData("linestyle.linestartarrowwidth"));
                AddProp(LINESTYLE__LINEESTARTARROWLength, GetData("linestyle.lineestartarrowlength"));
                AddProp(LINESTYLE__LINEENDARROWWIDTH, GetData("linestyle.lineendarrowwidth"));
                AddProp(LINESTYLE__LINEENDARROWLength, GetData("linestyle.lineendarrowlength"));
                AddProp(LINESTYLE__LINEJOINSTYLE, GetData("linestyle.linejoinstyle"));
                AddProp(LINESTYLE__LINEENDCAPSTYLE, GetData("linestyle.lineendcapstyle"));
                AddProp(LINESTYLE__ARROWHEADSOK, GetData("linestyle.arrowheadsok"));
                AddProp(LINESTYLE__ANYLINE, GetData("linestyle.anyline"));
                AddProp(LINESTYLE__HITLINETEST, GetData("linestyle.hitlinetest"));
                AddProp(LINESTYLE__LINEFILLSHAPE, GetData("linestyle.linefillshape"));
                AddProp(LINESTYLE__NOLINEDRAWDASH, GetData("linestyle.nolinedrawdash", EscherPropertyMetaData.TYPE_bool));
                AddProp(SHADOWSTYLE__TYPE, GetData("shadowstyle.type"));
                AddProp(SHADOWSTYLE__COLOR, GetData("shadowstyle.color", EscherPropertyMetaData.TYPE_RGB));
                AddProp(SHADOWSTYLE__HIGHLIGHT, GetData("shadowstyle.highlight"));
                AddProp(SHADOWSTYLE__CRMOD, GetData("shadowstyle.crmod"));
                AddProp(SHADOWSTYLE__OPACITY, GetData("shadowstyle.opacity"));
                AddProp(SHADOWSTYLE__OFFSetX, GetData("shadowstyle.offsetx"));
                AddProp(SHADOWSTYLE__OFFSetY, GetData("shadowstyle.offsety"));
                AddProp(SHADOWSTYLE__SECONDOFFSetX, GetData("shadowstyle.secondoffsetx"));
                AddProp(SHADOWSTYLE__SECONDOFFSetY, GetData("shadowstyle.secondoffsety"));
                AddProp(SHADOWSTYLE__SCALEXTOX, GetData("shadowstyle.scalextox"));
                AddProp(SHADOWSTYLE__SCALEYTOX, GetData("shadowstyle.scaleytox"));
                AddProp(SHADOWSTYLE__SCALEXTOY, GetData("shadowstyle.scalextoy"));
                AddProp(SHADOWSTYLE__SCALEYTOY, GetData("shadowstyle.scaleytoy"));
                AddProp(SHADOWSTYLE__PERSPECTIVEX, GetData("shadowstyle.perspectivex"));
                AddProp(SHADOWSTYLE__PERSPECTIVEY, GetData("shadowstyle.perspectivey"));
                AddProp(SHADOWSTYLE__WEIGHT, GetData("shadowstyle.weight"));
                AddProp(SHADOWSTYLE__ORIGINX, GetData("shadowstyle.originx"));
                AddProp(SHADOWSTYLE__ORIGINY, GetData("shadowstyle.originy"));
                AddProp(SHADOWSTYLE__SHADOW, GetData("shadowstyle.shadow"));
                AddProp(SHADOWSTYLE__SHADOWOBSURED, GetData("shadowstyle.shadowobsured"));
                AddProp(PERSPECTIVE__TYPE, GetData("perspective.type"));
                AddProp(PERSPECTIVE__OFFSetX, GetData("perspective.offsetx"));
                AddProp(PERSPECTIVE__OFFSetY, GetData("perspective.offsety"));
                AddProp(PERSPECTIVE__SCALEXTOX, GetData("perspective.scalextox"));
                AddProp(PERSPECTIVE__SCALEYTOX, GetData("perspective.scaleytox"));
                AddProp(PERSPECTIVE__SCALEXTOY, GetData("perspective.scalextoy"));
                AddProp(PERSPECTIVE__SCALEYTOY, GetData("perspective.scaleytoy"));
                AddProp(PERSPECTIVE__PERSPECTIVEX, GetData("perspective.perspectivex"));
                AddProp(PERSPECTIVE__PERSPECTIVEY, GetData("perspective.perspectivey"));
                AddProp(PERSPECTIVE__WEIGHT, GetData("perspective.weight"));
                AddProp(PERSPECTIVE__ORIGINX, GetData("perspective.originx"));
                AddProp(PERSPECTIVE__ORIGINY, GetData("perspective.originy"));
                AddProp(PERSPECTIVE__PERSPECTIVEON, GetData("perspective.perspectiveon"));
                AddProp(THREED__SPECULARAMOUNT, GetData("3d.specularamount"));
                AddProp(THREED__DIFFUSEAMOUNT, GetData("3d.diffuseamount"));
                AddProp(THREED__SHININESS, GetData("3d.shininess"));
                AddProp(THREED__EDGetHICKNESS, GetData("3d.edGethickness"));
                AddProp(THREED__EXTRUDEFORWARD, GetData("3d.extrudeforward"));
                AddProp(THREED__EXTRUDEBACKWARD, GetData("3d.extrudebackward"));
                AddProp(THREED__EXTRUDEPLANE, GetData("3d.extrudeplane"));
                AddProp(THREED__EXTRUSIONCOLOR, GetData("3d.extrusioncolor", EscherPropertyMetaData.TYPE_RGB));
                AddProp(THREED__CRMOD, GetData("3d.crmod"));
                AddProp(THREED__3DEFFECT, GetData("3d.3deffect"));
                AddProp(THREED__METALLIC, GetData("3d.metallic"));
                AddProp(THREED__USEEXTRUSIONCOLOR, GetData("3d.useextrusioncolor", EscherPropertyMetaData.TYPE_RGB));
                AddProp(THREED__LIGHTFACE, GetData("3d.lightface"));
                AddProp(THREEDSTYLE__YROTATIONANGLE, GetData("3dstyle.yrotationangle"));
                AddProp(THREEDSTYLE__XROTATIONANGLE, GetData("3dstyle.xrotationangle"));
                AddProp(THREEDSTYLE__ROTATIONAXISX, GetData("3dstyle.rotationaxisx"));
                AddProp(THREEDSTYLE__ROTATIONAXISY, GetData("3dstyle.rotationaxisy"));
                AddProp(THREEDSTYLE__ROTATIONAXISZ, GetData("3dstyle.rotationaxisz"));
                AddProp(THREEDSTYLE__ROTATIONANGLE, GetData("3dstyle.rotationangle"));
                AddProp(THREEDSTYLE__ROTATIONCENTERX, GetData("3dstyle.rotationcenterx"));
                AddProp(THREEDSTYLE__ROTATIONCENTERY, GetData("3dstyle.rotationcentery"));
                AddProp(THREEDSTYLE__ROTATIONCENTERZ, GetData("3dstyle.rotationcenterz"));
                AddProp(THREEDSTYLE__RENDERMODE, GetData("3dstyle.rendermode"));
                AddProp(THREEDSTYLE__TOLERANCE, GetData("3dstyle.tolerance"));
                AddProp(THREEDSTYLE__XVIEWPOINT, GetData("3dstyle.xviewpoint"));
                AddProp(THREEDSTYLE__YVIEWPOINT, GetData("3dstyle.yviewpoint"));
                AddProp(THREEDSTYLE__ZVIEWPOINT, GetData("3dstyle.zviewpoint"));
                AddProp(THREEDSTYLE__ORIGINX, GetData("3dstyle.originx"));
                AddProp(THREEDSTYLE__ORIGINY, GetData("3dstyle.originy"));
                AddProp(THREEDSTYLE__SKEWANGLE, GetData("3dstyle.skewangle"));
                AddProp(THREEDSTYLE__SKEWAMOUNT, GetData("3dstyle.skewamount"));
                AddProp(THREEDSTYLE__AMBIENTINTENSITY, GetData("3dstyle.ambientintensity"));
                AddProp(THREEDSTYLE__KEYX, GetData("3dstyle.keyx"));
                AddProp(THREEDSTYLE__KEYY, GetData("3dstyle.keyy"));
                AddProp(THREEDSTYLE__KEYZ, GetData("3dstyle.keyz"));
                AddProp(THREEDSTYLE__KEYINTENSITY, GetData("3dstyle.keyintensity"));
                AddProp(THREEDSTYLE__FillX, GetData("3dstyle.fillx"));
                AddProp(THREEDSTYLE__FillY, GetData("3dstyle.filly"));
                AddProp(THREEDSTYLE__FillZ, GetData("3dstyle.fillz"));
                AddProp(THREEDSTYLE__FillINTENSITY, GetData("3dstyle.fillintensity"));
                AddProp(THREEDSTYLE__CONSTRAINROTATION, GetData("3dstyle.constrainrotation"));
                AddProp(THREEDSTYLE__ROTATIONCENTERAUTO, GetData("3dstyle.rotationcenterauto"));
                AddProp(THREEDSTYLE__PARALLEL, GetData("3dstyle.parallel"));
                AddProp(THREEDSTYLE__KEYHARSH, GetData("3dstyle.keyharsh"));
                AddProp(THREEDSTYLE__FillHARSH, GetData("3dstyle.fillharsh"));
                AddProp(SHAPE__MASTER, GetData("shape.master"));
                AddProp(SHAPE__CONNECTORSTYLE, GetData("shape.connectorstyle"));
                AddProp(SHAPE__BLACKANDWHITESetTINGS, GetData("shape.blackandwhiteSettings"));
                AddProp(SHAPE__WMODEPUREBW, GetData("shape.wmodepurebw"));
                AddProp(SHAPE__WMODEBW, GetData("shape.wmodebw"));
                AddProp(SHAPE__OLEICON, GetData("shape.oleicon"));
                AddProp(SHAPE__PREFERRELATIVERESIZE, GetData("shape.preferrelativeresize"));
                AddProp(SHAPE__LOCKSHAPETYPE, GetData("shape.lockshapetype"));
                AddProp(SHAPE__DELETEATTACHEDOBJECT, GetData("shape.deleteattachedobject"));
                AddProp(SHAPE__BACKGROUNDSHAPE, GetData("shape.backgroundshape"));
                AddProp(CALLOUT__CALLOUTTYPE, GetData("callout.callouttype"));
                AddProp(CALLOUT__XYCALLOUTGAP, GetData("callout.xycalloutgap"));
                AddProp(CALLOUT__CALLOUTANGLE, GetData("callout.calloutangle"));
                AddProp(CALLOUT__CALLOUTDROPTYPE, GetData("callout.calloutdroptype"));
                AddProp(CALLOUT__CALLOUTDROPSPECIFIED, GetData("callout.calloutdropspecified"));
                AddProp(CALLOUT__CALLOUTLENGTHSPECIFIED, GetData("callout.calloutlengthspecified"));
                AddProp(CALLOUT__ISCALLOUT, GetData("callout.iscallout"));
                AddProp(CALLOUT__CALLOUTACCENTBAR, GetData("callout.calloutaccentbar"));
                AddProp(CALLOUT__CALLOUTTEXTBORDER, GetData("callout.callouttextborder"));
                AddProp(CALLOUT__CALLOUTMINUSX, GetData("callout.calloutminusx"));
                AddProp(CALLOUT__CALLOUTMINUSY, GetData("callout.calloutminusy"));
                AddProp(CALLOUT__DROPAUTO, GetData("callout.dropauto"));
                AddProp(CALLOUT__LENGTHSPECIFIED, GetData("callout.lengthspecified"));
                AddProp(GROUPSHAPE__SHAPENAME, GetData("groupshape.shapename"));
                AddProp(GROUPSHAPE__DESCRIPTION, GetData("groupshape.description"));
                AddProp(GROUPSHAPE__HYPERLINK, GetData("groupshape.hyperlink"));
                AddProp(GROUPSHAPE__WRAPPOLYGONVERTICES, GetData("groupshape.wrappolygonvertices", EscherPropertyMetaData.TYPE_ARRAY));
                AddProp(GROUPSHAPE__WRAPDISTLEFT, GetData("groupshape.wrapdistleft"));
                AddProp(GROUPSHAPE__WRAPDISTTOP, GetData("groupshape.wrapdisttop"));
                AddProp(GROUPSHAPE__WRAPDISTRIGHT, GetData("groupshape.wrapdistright"));
                AddProp(GROUPSHAPE__WRAPDISTBOTTOM, GetData("groupshape.wrapdistbottom"));
                AddProp(GROUPSHAPE__REGROUPID, GetData("groupshape.regroupid"));
                AddProp(GROUPSHAPE__EDITEDWRAP, GetData("groupshape.editedwrap"));
                AddProp(GROUPSHAPE__BEHINDDOCUMENT, GetData("groupshape.behinddocument"));
                AddProp(GROUPSHAPE__ONDBLCLICKNOTIFY, GetData("groupshape.ondblclicknotify"));
                AddProp(GROUPSHAPE__ISBUTTON, GetData("groupshape.isbutton"));
                AddProp(GROUPSHAPE__1DADJUSTMENT, GetData("groupshape.1dadjustment"));
                AddProp(GROUPSHAPE__HIDDEN, GetData("groupshape.hidden"));
                AddProp(GROUPSHAPE__PRINT, GetData("groupshape.print", EscherPropertyMetaData.TYPE_bool));
            }
        }

        /// <summary>
        /// Adds the prop.
        /// </summary>
        /// <param name="s">The s.</param>
        /// <param name="data">The data.</param>
        private static void AddProp(int s, EscherPropertyMetaData data)
        {
            properties[(short)s]= data;
        }


        /// <summary>
        /// Gets the data.
        /// </summary>
        /// <param name="propName">Name of the prop.</param>
        /// <param name="type">The type.</param>
        /// <returns></returns>
        private static EscherPropertyMetaData GetData(String propName, byte type)
        {
            return new EscherPropertyMetaData(propName, type);
        }


        /// <summary>
        /// Gets the data.
        /// </summary>
        /// <param name="propName">Name of the prop.</param>
        /// <returns></returns>
        private static EscherPropertyMetaData GetData(String propName)
        {
            return new EscherPropertyMetaData(propName);
        }

        /// <summary>
        /// Gets the name of the property.
        /// </summary>
        /// <param name="propertyId">The property id.</param>
        /// <returns></returns>
        public static String GetPropertyName(short propertyId)
        {
            InitProps();
            EscherPropertyMetaData o = (EscherPropertyMetaData)properties[propertyId];
            return o == null ? "unknown" : o.Description;
        }

        /// <summary>
        /// Gets the type of the property.
        /// </summary>
        /// <param name="propertyId">The property id.</param>
        /// <returns></returns>
        public static byte GetPropertyType(short propertyId)
        {
            InitProps();
            EscherPropertyMetaData escherPropertyMetaData = (EscherPropertyMetaData)properties[propertyId];
            return escherPropertyMetaData == null ? (byte)0 : escherPropertyMetaData.Type;
        }
    }
}


