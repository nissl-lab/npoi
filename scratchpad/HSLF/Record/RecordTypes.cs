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

namespace NPOI.HSLF.Record
{
    using System;
    using System.Reflection;
    using System.Collections;

    /**
     * List of all known record types in a PowerPoint document, and the
     *  classes that handle them.
     * There are two categories of records:
     * <li> PowerPoint records: 0 <= info <= 10002 (will carry class info)
     * <li> Escher records: info >= 0xF000 (handled by DDF, so no class info)
     *
     * @author Yegor Kozlov
     * @author Nick Burch
     */
    public class RecordTypes
    {
        public static Hashtable typeToName;
        public static Hashtable typeToClass;

        public static RecordType Unknown = new RecordType(0, null);
        //public static RecordType Document = new RecordType(1000, typeof(Document));
        //public static RecordType DocumentAtom = new RecordType(1001, typeof(DocumentAtom));
        //public static RecordType EndDocument = new RecordType(1002, null);
        //public static RecordType Slide = new RecordType(1006, typeof(Slide));
        //public static RecordType SlideAtom = new RecordType(1007, typeof(SlideAtom));
        //public static RecordType Notes = new RecordType(1008, typeof(Notes));
        //public static RecordType NotesAtom = new RecordType(1009, typeof(NotesAtom));
        //public static RecordType Environment = new RecordType(1010, typeof(Environment));
        //public static RecordType SlidePersistAtom = new RecordType(1011, typeof(SlidePersistAtom));
        //public static RecordType SSlideLayoutAtom = new RecordType(1015, null);
        //public static RecordType MainMaster = new RecordType(1016, typeof(MainMaster));
        //public static RecordType SSSlideInfoAtom = new RecordType(1017, null);
        //public static RecordType SlideViewInfo = new RecordType(1018, null);
        //public static RecordType GuideAtom = new RecordType(1019, null);
        //public static RecordType ViewInfo = new RecordType(1020, null);
        //public static RecordType ViewInfoAtom = new RecordType(1021, null);
        //public static RecordType SlideViewInfoAtom = new RecordType(1022, null);
        //public static RecordType VBAInfo = new RecordType(1023, null);
        //public static RecordType VBAInfoAtom = new RecordType(1024, null);
        //public static RecordType SSDocInfoAtom = new RecordType(1025, null);
        //public static RecordType Summary = new RecordType(1026, null);
        //public static RecordType DocRoutingSlip = new RecordType(1030, null);
        //public static RecordType OutlineViewInfo = new RecordType(1031, null);
        //public static RecordType SorterViewInfo = new RecordType(1032, null);
        //public static RecordType ExObjList = new RecordType(1033, typeof(ExObjList));
        //public static RecordType ExObjListAtom = new RecordType(1034, typeof(ExObjListAtom));
        //public static RecordType PPDrawingGroup = new RecordType(1035, typeof(PPDrawingGroup));
        //public static RecordType PPDrawing = new RecordType(1036, typeof(PPDrawing));
        //public static RecordType NamedShows = new RecordType(1040, null);
        //public static RecordType NamedShow = new RecordType(1041, null);
        //public static RecordType NamedShowSlides = new RecordType(1042, null);
        //public static RecordType SheetProperties = new RecordType(1044, null);
        //public static RecordType List = new RecordType(2000, null);
        //public static RecordType FontCollection = new RecordType(2005, typeof(FontCollection));
        //public static RecordType BookmarkCollection = new RecordType(2019, null);
        //public static RecordType SoundCollection = new RecordType(2020, typeof(SoundCollection));
        //public static RecordType SoundCollAtom = new RecordType(2021, null);
        //public static RecordType Sound = new RecordType(2022, typeof(Sound));
        public static RecordType SoundData = new RecordType(2023, typeof(SoundData));
        //public static RecordType BookmarkSeedAtom = new RecordType(2025, null);
        //public static RecordType ColorSchemeAtom = new RecordType(2032, typeof(ColorSchemeAtom));
        //public static RecordType ExObjRefAtom = new RecordType(3009, null);
        //public static RecordType OEShapeAtom = new RecordType(3009, typeof(OEShapeAtom));
        //public static RecordType OEPlaceholderAtom = new RecordType(3011, typeof(OEPlaceholderAtom));
        //public static RecordType GPopublicintAtom = new RecordType(3024, null);
        //public static RecordType GRatioAtom = new RecordType(3031, null);
        //public static RecordType OutlineTextRefAtom = new RecordType(3998, typeof(OutlineTextRefAtom));
        //public static RecordType TextHeaderAtom = new RecordType(3999, typeof(TextHeaderAtom));
        //public static RecordType TextCharsAtom = new RecordType(4000, typeof(TextCharsAtom));
        //public static RecordType StyleTextPropAtom = new RecordType(4001, typeof(StyleTextPropAtom));
        //public static RecordType BaseTextPropAtom = new RecordType(4002, null);
        //public static RecordType TxMasterStyleAtom = new RecordType(4003, typeof(TxMasterStyleAtom));
        //public static RecordType TxCFStyleAtom = new RecordType(4004, null);
        //public static RecordType TxPFStyleAtom = new RecordType(4005, null);
        public static RecordType TextRulerAtom = new RecordType(4006, typeof(TextRulerAtom));
        //public static RecordType TextBookmarkAtom = new RecordType(4007, null);
        //public static RecordType TextBytesAtom = new RecordType(4008, typeof(TextBytesAtom));
        //public static RecordType TxSIStyleAtom = new RecordType(4009, null);
        public static RecordType TextSpecInfoAtom = new RecordType(4010, typeof(TextSpecInfoAtom));
        //public static RecordType DefaultRulerAtom = new RecordType(4011, null);
        //public static RecordType FontEntityAtom = new RecordType(4023, typeof(FontEntityAtom));
        //public static RecordType FontEmbeddedData = new RecordType(4024, null);
        //public static RecordType CString = new RecordType(4026, typeof(CString));
        //public static RecordType MetaFile = new RecordType(4033, null);
        //public static RecordType ExOleObjAtom = new RecordType(4035, typeof(ExOleObjAtom));
        //public static RecordType SrKinsoku = new RecordType(4040, null);
        //public static RecordType HandOut = new RecordType(4041, typeof(DummyPositionSensitiveRecordWithChildren));
        //public static RecordType ExEmbed = new RecordType(4044, typeof(ExEmbed));
        //public static RecordType ExEmbedAtom = new RecordType(4045, typeof(ExEmbedAtom));
        //public static RecordType ExLink = new RecordType(4046, null);
        //public static RecordType BookmarkEntityAtom = new RecordType(4048, null);
        //public static RecordType ExLinkAtom = new RecordType(4049, null);
        //public static RecordType SrKinsokuAtom = new RecordType(4050, null);
        //public static RecordType ExHyperlinkAtom = new RecordType(4051, typeof(ExHyperlinkAtom));
        //public static RecordType ExHyperlink = new RecordType(4055, typeof(ExHyperlink));
        //public static RecordType SlideNumberMCAtom = new RecordType(4056, null);
        //public static RecordType HeadersFooters = new RecordType(4057, typeof(HeadersFootersContainer));
        //public static RecordType HeadersFootersAtom = new RecordType(4058, typeof(HeadersFootersAtom));
        public static RecordType TxInteractiveInfoAtom = new RecordType(4063, typeof(TxInteractiveInfoAtom));
        //public static RecordType CharFormatAtom = new RecordType(4066, null);
        //public static RecordType ParaFormatAtom = new RecordType(4067, null);
        //public static RecordType RecolorInfoAtom = new RecordType(4071, null);
        //public static RecordType ExQuickTimeMovie = new RecordType(4074, null);
        //public static RecordType ExQuickTimeMovieData = new RecordType(4075, null);
        //public static RecordType ExControl = new RecordType(4078, typeof(ExControl));
        //public static RecordType SlideListWithText = new RecordType(4080, typeof(SlideListWithText));
        //public static RecordType InteractiveInfo = new RecordType(4082, typeof(InteractiveInfo));
        //public static RecordType InteractiveInfoAtom = new RecordType(4083, typeof(InteractiveInfoAtom));
        //public static RecordType UserEditAtom = new RecordType(4085, typeof(UserEditAtom));
        //public static RecordType CurrentUserAtom = new RecordType(4086, null);
        //public static RecordType DateTimeMCAtom = new RecordType(4087, null);
        //public static RecordType GenericDateMCAtom = new RecordType(4088, null);
        //public static RecordType FooterMCAtom = new RecordType(4090, null);
        //public static RecordType ExControlAtom = new RecordType(4091, typeof(ExControlAtom));
        //public static RecordType ExMediaAtom = new RecordType(4100, typeof(ExMediaAtom));
        //public static RecordType ExVideoContainer = new RecordType(4101, typeof(ExVideoContainer));
        //public static RecordType ExAviMovie = new RecordType(4102, typeof(ExAviMovie));
        //public static RecordType ExMCIMovie = new RecordType(4103, typeof(ExMCIMovie));
        //public static RecordType ExMIDIAudio = new RecordType(4109, null);
        //public static RecordType ExCDAudio = new RecordType(4110, null);
        //public static RecordType ExWAVAudioEmbedded = new RecordType(4111, null);
        //public static RecordType ExWAVAudioLink = new RecordType(4112, null);
        //public static RecordType ExOleObjStg = new RecordType(4113, typeof(ExOleObjStg));
        //public static RecordType ExCDAudioAtom = new RecordType(4114, null);
        //public static RecordType ExWAVAudioEmbeddedAtom = new RecordType(4115, null);
        public static RecordType AnimationInfo = new RecordType(4116, typeof(AnimationInfo));
        public static RecordType AnimationInfoAtom = new RecordType(4081, typeof(AnimationInfoAtom));
        //public static RecordType RTFDateTimeMCAtom = new RecordType(4117, null);
        //public static RecordType ProgTags = new RecordType(5000, typeof(DummyPositionSensitiveRecordWithChildren));
        //public static RecordType ProgStringTag = new RecordType(5001, null);
        //public static RecordType ProgBinaryTag = new RecordType(5002, typeof(DummyPositionSensitiveRecordWithChildren));
        //public static RecordType BinaryTagData = new RecordType(5003, typeof(DummyPositionSensitiveRecordWithChildren));
        //public static RecordType PrpublicintOptions = new RecordType(6000, null);
        //public static RecordType PersistPtrFullBlock = new RecordType(6001, typeof(PersistPtrHolder));
        //public static RecordType PersistPtrIncrementalBlock = new RecordType(6002, typeof(PersistPtrHolder));
        //public static RecordType GScalingAtom = new RecordType(10001, null);
        //public static RecordType GRColorAtom = new RecordType(10002, null);

        //// Records ~12000 seem to be related to the Comments used in PPT 2000/XP
        //// (Comments in PPT97 are normal Escher text boxes)
        //public static RecordType Comment2000 = new RecordType(12000, typeof(Comment2000));
        public static RecordType Comment2000Atom = new RecordType(12001, typeof(Comment2000Atom));
        //public static RecordType Comment2000Summary = new RecordType(12004, null);
        //public static RecordType Comment2000SummaryAtom = new RecordType(12005, null);

        //// Records ~12050 seem to be related to Document Encryption
        //public static RecordType DocumentEncryptionAtom = new RecordType(12052, typeof(DocumentEncryptionAtom));

        //public static RecordType OriginalMainMasterId = new RecordType(1052, null);
        //public static RecordType CompositeMasterId = new RecordType(1052, null);
        //public static RecordType RoundTripContentMasterInfo12 = new RecordType(1054, null);
        //public static RecordType RoundTripShapeId12 = new RecordType(1055, null);
        public static RecordType RoundTripHFPlaceholder12 = new RecordType(1056, typeof(RoundTripHFPlaceholder12));
        //public static RecordType RoundTripContentMasterId = new RecordType(1058, null);
        //public static RecordType RoundTripOArtTextStyles12 = new RecordType(1059, null);
        //public static RecordType RoundTripShapeCheckSumForCustomLayouts12 = new RecordType(1062, null);
        //public static RecordType RoundTripNotesMasterTextStyles12 = new RecordType(1063, null);
        //public static RecordType RoundTripCustomTableStyles12 = new RecordType(1064, null);

        //records greater then 0xF000 belong to with Microsoft Office Drawing format also known as Escher
        public static int EscherDggContainer = 0xf000;
        public static int EscherDgg = 0xf006;
        public static int EscherCLSID = 0xf016;
        public static int EscherOPT = 0xf00b;
        public static int EscherBStoreContainer = 0xf001;
        public static int EscherBSE = 0xf007;
        public static int EscherBlip_START = 0xf018;
        public static int EscherBlip_END = 0xf117;
        public static int EscherDgContainer = 0xf002;
        public static int EscherDg = 0xf008;
        public static int EscherRegroupItems = 0xf118;
        public static int EscherColorScheme = 0xf120;
        public static int EscherSpgrContainer = 0xf003;
        public static int EscherSpContainer = 0xf004;
        public static int EscherSpgr = 0xf009;
        public static int EscherSp = 0xf00a;
        public static int EscherTextbox = 0xf00c;
        public static int EscherClientTextbox = 0xf00d;
        public static int EscherAnchor = 0xf00e;
        public static int EscherChildAnchor = 0xf00f;
        public static int EscherClientAnchor = 0xf010;
        public static int EscherClientData = 0xf011;
        public static int EscherSolverContainer = 0xf005;
        public static int EscherConnectorRule = 0xf012;
        public static int EscherAlignRule = 0xf013;
        public static int EscherArcRule = 0xf014;
        public static int EscherClientRule = 0xf015;
        public static int EscherCalloutRule = 0xf017;
        public static int EscherSelection = 0xf119;
        public static int EscherColorMRU = 0xf11a;
        public static int EscherDeletedPspl = 0xf11d;
        public static int EscherSplitMenuColors = 0xf11e;
        public static int EscherOleObject = 0xf11f;
        public static int EscherUserDefined = 0xf122;

        /**
         * Returns name of the record by its type
         *
         * @param type section of the record header
         * @return name of the record
         */
        public static String RecordName(int type)
        {
            String name = (String)typeToName[type];
            if (name == null) name = "Unknown" + type;
            return name;
        }

        /**
         * Returns the class handling a record by its type.
         * If given an un-handled PowerPoint record, will return a dummy
         *  placeholder class. If given an unknown PowerPoint record, or
         *  and Escher record, will return null.
         *
         * @param type section of the record header
         * @return class to handle the record, or null if an unknown (eg Escher) record
         */
        public static Type RecordHandlingClass(int type)
        {
            Type c = (Type)typeToClass[type];
            return c;
        }

        static RecordTypes()
        {
            typeToName = new Hashtable();
            typeToClass = new Hashtable();
            try
            {
                FieldInfo[] f = typeof(RecordTypes).GetFields();
                for (int i = 0; i < f.Length; i++)
                {
                    Object val = f[i].GetType();

                    // Escher record, only store ID -> Name
                    if (val is Int32)
                    {
                        typeToName[val]= f[i].Name;
                    }
                    // PowerPoint record, store ID -> Name and ID -> Class
                    if (val is RecordType)
                    {
                        RecordType t = (RecordType)val;
                        Type c = t.HandlingClass;
                        int id = t.typeID;
                        if (c == null) { c = typeof(UnknownRecordPlaceholder); }

                        typeToName[id]= f[i].Name;
                        typeToClass[id]=c;
                    }
                }
            }
            catch (InvalidOperationException)
            {
                throw new Exception("Failed to Initialize records types");
            }
        }


        /**
         * Wrapper for the details of a PowerPoint or Escher record type.
         * Contains both the type, and the handling class (if any), and
         *  offers methods to get either back out.
         */
        public class RecordType
        {
            public int typeID;
            public Type HandlingClass;
            public RecordType(int typeID, Type handlingClass)
            {
                this.typeID = typeID;
                this.HandlingClass = handlingClass;
            }
        }
    }
}