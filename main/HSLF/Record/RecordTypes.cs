using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.Record
{
	public class RecordTypes
	{
		public static RecordTypes Unknown => RecordTypesLookup[0];
		public static RecordTypes UnknownRecordPlaceholder => RecordTypesLookup[-1];
		public static RecordTypes Document => RecordTypesLookup[1000];
		public static RecordTypes DocumentAtom => RecordTypesLookup[1001];
		public static RecordTypes EndDocument => RecordTypesLookup[1002];
		public static RecordTypes Slide => RecordTypesLookup[1006];
		public static RecordTypes SlideAtom => RecordTypesLookup[1007];
		public static RecordTypes Notes => RecordTypesLookup[1008];
		public static RecordTypes NotesAtom => RecordTypesLookup[1009];
		public static RecordTypes Environment => RecordTypesLookup[1010];
		public static RecordTypes SlidePersistAtom => RecordTypesLookup[1011];
		public static RecordTypes SSlideLayoutAtom => RecordTypesLookup[1015];
		public static RecordTypes MainMaster => RecordTypesLookup[1016];
		public static RecordTypes SSSlideInfoAtom => RecordTypesLookup[1017];
		public static RecordTypes SlideViewInfo => RecordTypesLookup[1018];
		public static RecordTypes GuideAtom => RecordTypesLookup[1019];
		public static RecordTypes ViewInfo => RecordTypesLookup[1020];
		public static RecordTypes ViewInfoAtom => RecordTypesLookup[1021];
		public static RecordTypes SlideViewInfoAtom => RecordTypesLookup[1022];
		public static RecordTypes VBAInfo => RecordTypesLookup[1023];
		public static RecordTypes VBAInfoAtom => RecordTypesLookup[1024];
		public static RecordTypes SSDocInfoAtom => RecordTypesLookup[1025];
		public static RecordTypes Summary => RecordTypesLookup[1026];
		public static RecordTypes DocRoutingSlip => RecordTypesLookup[1030];
		public static RecordTypes OutlineViewInfo => RecordTypesLookup[1031];
		public static RecordTypes SorterViewInfo => RecordTypesLookup[1032];
		public static RecordTypes ExObjList => RecordTypesLookup[1033];
		public static RecordTypes ExObjListAtom => RecordTypesLookup[1034];
		public static RecordTypes PPDrawingGroup => RecordTypesLookup[1035];
		public static RecordTypes PPDrawing => RecordTypesLookup[1036];
		public static RecordTypes NamedShows => RecordTypesLookup[1040];
		public static RecordTypes NamedShow => RecordTypesLookup[1041];
		public static RecordTypes NamedShowSlides => RecordTypesLookup[1042];
		public static RecordTypes SheetProperties => RecordTypesLookup[1044];
		public static RecordTypes OriginalMainMasterId => RecordTypesLookup[1052];
		public static RecordTypes CompositeMasterId => RecordTypesLookup[1052];
		public static RecordTypes RoundTripContentMasterInfo12 => RecordTypesLookup[1054];
		public static RecordTypes RoundTripShapeId12 => RecordTypesLookup[1055];
		public static RecordTypes RoundTripHFPlaceholder12 => RecordTypesLookup[1056];
		public static RecordTypes RoundTripContentMasterId => RecordTypesLookup[1058];
		public static RecordTypes RoundTripOArtTextStyles12 => RecordTypesLookup[1059];
		public static RecordTypes RoundTripShapeCheckSumForCustomLayouts12 => RecordTypesLookup[1062];
		public static RecordTypes RoundTripNotesMasterTextStyles12 => RecordTypesLookup[1063];
		public static RecordTypes RoundTripCustomTableStyles12 => RecordTypesLookup[1064];
		public static RecordTypes List => RecordTypesLookup[2000];
		public static RecordTypes FontCollection => RecordTypesLookup[2005];
		public static RecordTypes BookmarkCollection => RecordTypesLookup[2019];
		public static RecordTypes SoundCollection => RecordTypesLookup[2020];
		public static RecordTypes SoundCollAtom => RecordTypesLookup[2021];
		public static RecordTypes Sound => RecordTypesLookup[2022];
		public static RecordTypes SoundData => RecordTypesLookup[2023];
		public static RecordTypes BookmarkSeedAtom => RecordTypesLookup[2025];
		public static RecordTypes ColorSchemeAtom => RecordTypesLookup[2032];
		public static RecordTypes ExObjRefAtom => RecordTypesLookup[3009];
		public static RecordTypes OEPlaceholderAtom => RecordTypesLookup[3011];
		public static RecordTypes GPopublicintAtom => RecordTypesLookup[3024];
		public static RecordTypes GRatioAtom => RecordTypesLookup[3031];
		public static RecordTypes OutlineTextRefAtom => RecordTypesLookup[3998];
		public static RecordTypes TextHeaderAtom => RecordTypesLookup[3999];
		public static RecordTypes TextCharsAtom => RecordTypesLookup[4000];
		public static RecordTypes StyleTextPropAtom => RecordTypesLookup[4001];//0x0fa1 RT_StyleTextPropAtom
		public static RecordTypes MasterTextPropAtom => RecordTypesLookup[4002];
		public static RecordTypes TxMasterStyleAtom => RecordTypesLookup[4003];
		public static RecordTypes TxCFStyleAtom => RecordTypesLookup[4004];
		public static RecordTypes TxPFStyleAtom => RecordTypesLookup[4005];
		public static RecordTypes TextRulerAtom => RecordTypesLookup[4006];
		public static RecordTypes TextBookmarkAtom => RecordTypesLookup[4007];
		public static RecordTypes TextBytesAtom => RecordTypesLookup[4008];
		public static RecordTypes TxSIStyleAtom => RecordTypesLookup[4009];
		public static RecordTypes TextSpecInfoAtom => RecordTypesLookup[4010];
		public static RecordTypes DefaultRulerAtom => RecordTypesLookup[4011];
		public static RecordTypes StyleTextProp9Atom => RecordTypesLookup[4012]; //0x0FAC RT_StyleTextProp9Atom
		public static RecordTypes FontEntityAtom => RecordTypesLookup[4023];
		public static RecordTypes FontEmbeddedData => RecordTypesLookup[4024];
		public static RecordTypes CString => RecordTypesLookup[4026];
		public static RecordTypes MetaFile => RecordTypesLookup[4033];
		public static RecordTypes ExOleObjAtom => RecordTypesLookup[4035];
		public static RecordTypes SrKinsoku => RecordTypesLookup[4040];
		public static RecordTypes HandOut => RecordTypesLookup[4041];
		public static RecordTypes ExEmbed => RecordTypesLookup[4044];
		public static RecordTypes ExEmbedAtom => RecordTypesLookup[4045];
		public static RecordTypes ExLink => RecordTypesLookup[4046];
		public static RecordTypes BookmarkEntityAtom => RecordTypesLookup[4048];
		public static RecordTypes ExLinkAtom => RecordTypesLookup[4049];
		public static RecordTypes SrKinsokuAtom => RecordTypesLookup[4050];
		public static RecordTypes ExHyperlinkAtom => RecordTypesLookup[4051];
		public static RecordTypes ExHyperlink => RecordTypesLookup[4055];
		public static RecordTypes SlideNumberMCAtom => RecordTypesLookup[4056];
		public static RecordTypes HeadersFooters => RecordTypesLookup[4057];
		public static RecordTypes HeadersFootersAtom => RecordTypesLookup[4058];
		public static RecordTypes TxInteractiveInfoAtom => RecordTypesLookup[4063];
		public static RecordTypes CharFormatAtom => RecordTypesLookup[4066];
		public static RecordTypes ParaFormatAtom => RecordTypesLookup[4067];
		public static RecordTypes RecolorInfoAtom => RecordTypesLookup[4071];
		public static RecordTypes ExQuickTimeMovie => RecordTypesLookup[4074];
		public static RecordTypes ExQuickTimeMovieData => RecordTypesLookup[4075];
		public static RecordTypes ExControl => RecordTypesLookup[4078];
		public static RecordTypes SlideListWithText => RecordTypesLookup[4080];
		public static RecordTypes InteractiveInfo => RecordTypesLookup[4082];
		public static RecordTypes InteractiveInfoAtom => RecordTypesLookup[4083];
		public static RecordTypes UserEditAtom => RecordTypesLookup[4085];
		public static RecordTypes CurrentUserAtom => RecordTypesLookup[4086];
		public static RecordTypes DateTimeMCAtom => RecordTypesLookup[4087];
		public static RecordTypes GenericDateMCAtom => RecordTypesLookup[4088];
		public static RecordTypes FooterMCAtom => RecordTypesLookup[4090];
		public static RecordTypes ExControlAtom => RecordTypesLookup[4091];
		public static RecordTypes ExMediaAtom => RecordTypesLookup[4100];
		public static RecordTypes ExVideoContainer => RecordTypesLookup[4101];
		public static RecordTypes ExAviMovie => RecordTypesLookup[4102];
		public static RecordTypes ExMCIMovie => RecordTypesLookup[4103];
		public static RecordTypes ExMIDIAudio => RecordTypesLookup[4109];
		public static RecordTypes ExCDAudio => RecordTypesLookup[4110];
		public static RecordTypes ExWAVAudioEmbedded => RecordTypesLookup[4111];
		public static RecordTypes ExWAVAudioLink => RecordTypesLookup[4112];
		public static RecordTypes ExOleObjStg => RecordTypesLookup[4113];
		public static RecordTypes ExCDAudioAtom => RecordTypesLookup[4114];
		public static RecordTypes ExWAVAudioEmbeddedAtom => RecordTypesLookup[4115];
		public static RecordTypes AnimationInfo => RecordTypesLookup[4116];
		public static RecordTypes AnimationInfoAtom => RecordTypesLookup[4081];
		public static RecordTypes RTFDateTimeMCAtom => RecordTypesLookup[4117];
		public static RecordTypes ProgTags => RecordTypesLookup[5000];
		public static RecordTypes ProgStringTag => RecordTypesLookup[5001];
		public static RecordTypes ProgBinaryTag => RecordTypesLookup[5002];
		public static RecordTypes BinaryTagData => RecordTypesLookup[5003];//0x138b RT_BinaryTagDataBlob
		public static RecordTypes PrpublicintOptions => RecordTypesLookup[6000];
		public static RecordTypes PersistPtrFullBlock => RecordTypesLookup[6001];
		public static RecordTypes PersistPtrIncrementalBlock => RecordTypesLookup[6002];
		public static RecordTypes GScalingAtom => RecordTypesLookup[10001];
		public static RecordTypes GRColorAtom => RecordTypesLookup[10002];

		// Records ~12000 seem to be related to the Comments used in PPT 2000/XP
		// (Comments in PPT97 are normal Escher text boxes)
		public static RecordTypes Comment2000 => RecordTypesLookup[12000];
		public static RecordTypes Comment2000Atom => RecordTypesLookup[12001];
		public static RecordTypes Comment2000Summary => RecordTypesLookup[12004];
		public static RecordTypes Comment2000SummaryAtom => RecordTypesLookup[12005];

		// Records ~12050 seem to be related to Document Encryption
		public static RecordTypes DocumentEncryptionAtom => RecordTypesLookup[12052];

		public static readonly Dictionary<int, RecordTypes> RecordTypesLookup =
			new Dictionary<int, RecordTypes>()
			{
				{ 0, new RecordTypes(0, null) },
				{ -1, new RecordTypes(-1, HSLF.Record.UnknownRecordPlaceholder.GetInstance) },
				{1000, new RecordTypes(1000, new Document) },
				{1001, new RecordTypes(1001, new DocumentAtom) },
				{1002, new RecordTypes(1002, null) },
				{1006, new RecordTypes(1006, new Slide) },
				{1007, new RecordTypes(1007, new SlideAtom) },
				{1008, new RecordTypes(1008, new Notes) },
				{1009, new RecordTypes(1009, new NotesAtom) },
				{1010, new RecordTypes(1010, new Environment) },
				{1011, new RecordTypes(1011, new SlidePersistAtom) },
				{1015, new RecordTypes(1015,null) },
				{1016, new RecordTypes(1016, new MainMaster) },
				{1017, new RecordTypes(1017, new SSSlideInfoAtom) },
				{1018, new RecordTypes(1018,null) },
				{1019, new RecordTypes(1019,null) },
				{1020, new RecordTypes(1020,null) },
				{1021, new RecordTypes(1021,null) },
				{1022, new RecordTypes(1022,null) },
				{1023, new RecordTypes(1023, new VBAInfoContainer) },
				{1024, new RecordTypes(1024, new VBAInfoAtom) },
				{1025, new RecordTypes(1025,null) },
				{1026, new RecordTypes(1026,null) },
				{1030, new RecordTypes(1030,null) },
				{1031, new RecordTypes(1031,null) },
				{1032, new RecordTypes(1032,null) },
				{1033, new RecordTypes(1033, new ExObjList) },
				{1034, new RecordTypes(1034, new ExObjListAtom) },
				{1035, new RecordTypes(1035, new PPDrawingGroup) },
				{1036, new RecordTypes(1036, new PPDrawing) },
				{1040, new RecordTypes(1040,null) },
				{1041, new RecordTypes(1041,null) },
				{1042, new RecordTypes(1042,null) },
				{1044, new RecordTypes(1044,null) },
				{1052, new RecordTypes(1052,null) },
				{1052, new RecordTypes(1052,null) },
				{1054, new RecordTypes(1054,null) },
				{1055, new RecordTypes(1055,null) },
				{1056, new RecordTypes(1056, new RoundTripHFPlaceholder12) },
				{1058, new RecordTypes(1058,null) },
				{1059, new RecordTypes(1059,null) },
				{1062, new RecordTypes(1062,null) },
				{1063, new RecordTypes(1063,null) },
				{1064, new RecordTypes(1064,null) },

				{2000, new RecordTypes(2000, new DocInfoListContainer) },
				{2005, new RecordTypes(2005, new FontCollection) },
				{2019, new RecordTypes(2019,null) },
				{2020, new RecordTypes(2020, new SoundCollection) },
				{2021, new RecordTypes(2021,null) },
				{2022, new RecordTypes(2022, new Sound) },
				{2023, new RecordTypes(2023, new SoundData) },
				{2025, new RecordTypes(2025,null) },
				{2032, new RecordTypes(2032, new ColorSchemeAtom) },
				{3009, new RecordTypes(3009,new ExObjRefAtom) },
				{3011, new RecordTypes(3011, new OEPlaceholderAtom) },
				{3024, new RecordTypes(3024,null) },
				{3031, new RecordTypes(3031,null) },
				{3998, new RecordTypes(3998, new OutlineTextRefAtom) },
				{3999, new RecordTypes(3999, new TextHeaderAtom) },
				{4000, new RecordTypes(4000, new TextCharsAtom) },
				{4001, new RecordTypes(4001, new StyleTextPropAtom) },//0x0fa1 RT_StyleTextPropAtom
				{4002, new RecordTypes(4002, new MasterTextPropAtom) },
				{4003, new RecordTypes(4003, new TxMasterStyleAtom) },
				{4004, new RecordTypes(4004,null) },
				{4005, new RecordTypes(4005,null) },
				{4006, new RecordTypes(4006, new TextRulerAtom) },
				{4007, new RecordTypes(4007,null) },
				{4008, new RecordTypes(4008, new TextBytesAtom) },
				{4009, new RecordTypes(4009,null) },
				{4010, new RecordTypes(4010, new TextSpecInfoAtom) },
				{4011, new RecordTypes(4011,null) },
				{4012, new RecordTypes(4012, new StyleTextProp9Atom) }, //0x0FAC RT_StyleTextProp9Atom
				{4023, new RecordTypes(4023, new FontEntityAtom) },
				{4024, new RecordTypes(4024, new FontEmbeddedData) },
				{4026, new RecordTypes(4026, new CString()) },
				{4033, new RecordTypes(4033,null) },
				{4035, new RecordTypes(4035, new ExOleObjAtom) },
				{4040, new RecordTypes(4040,null) },
				{4041, new RecordTypes(4041, new DummyPositionSensitiveRecordWithChildren) },
				{4044, new RecordTypes(4044, new ExEmbed) },
				{4045, new RecordTypes(4045, new ExEmbedAtom) },
				{4046, new RecordTypes(4046,null) },
				{4048, new RecordTypes(4048,null) },
				{4049, new RecordTypes(4049,null) },
				{4050, new RecordTypes(4050,null) },
				{4051, new RecordTypes(4051, new ExHyperlinkAtom) },
				{4055, new RecordTypes(4055, new ExHyperlink) },
				{4056, new RecordTypes(4056,null) },
				{4057, new RecordTypes(4057, new HeadersFootersContainer) },
				{4058, new RecordTypes(4058, new HeadersFootersAtom) },
				{4063, new RecordTypes(4063, new TxInteractiveInfoAtom) },
				{4066, new RecordTypes(4066,null) },
				{4067, new RecordTypes(4067,null) },
				{4071, new RecordTypes(4071,null) },
				{4074, new RecordTypes(4074,null) },
				{4075, new RecordTypes(4075,null) },
				{4078, new RecordTypes(4078, new ExControl) },
				{4080, new RecordTypes(4080, new SlideListWithText) },
				{4082, new RecordTypes(4082, new InteractiveInfo) },
				{4083, new RecordTypes(4083, new InteractiveInfoAtom) },
				{4085, new RecordTypes(4085, new UserEditAtom) },
				{4086, new RecordTypes(4086,null) },
				{4087, new RecordTypes(4087, new DateTimeMCAtom) },
				{4088, new RecordTypes(4088,null) },
				{4090, new RecordTypes(4090,null) },
				{4091, new RecordTypes(4091, new ExControlAtom) },
				{4100, new RecordTypes(4100, new ExMediaAtom) },
				{4101, new RecordTypes(4101, new ExVideoContainer) },
				{4102, new RecordTypes(4102, new ExAviMovie) },
				{4103, new RecordTypes(4103, new ExMCIMovie) },
				{4109, new RecordTypes(4109,null) },
				{4110, new RecordTypes (4110,null) },
				{4111, new RecordTypes(4111,null) },
				{4112, new RecordTypes(4112,null) },
				{4113, new RecordTypes(4113, new ExOleObjStg) },
				{4114, new RecordTypes(4114,null) },
				{4115, new RecordTypes(4115,null) },
				{4116, new RecordTypes(4116, new AnimationInfo) },
				{4081, new RecordTypes(4081, new AnimationInfoAtom) },
				{4117, new RecordTypes(4117,null) },
				{5000, new RecordTypes(5000, new DummyPositionSensitiveRecordWithChildren) },
				{5001, new RecordTypes(5001,null) },
				{5002, new RecordTypes(5002, new DummyPositionSensitiveRecordWithChildren) },
				{5003, new RecordTypes(5003, new BinaryTagDataBlob) },//0x138b RT_BinaryTagDataBlob
				{6000, new RecordTypes(6000,null) },
				{6001, new RecordTypes(6001, new PersistPtrHolder) },
				{6002, new RecordTypes(6002, new PersistPtrHolder) },
				{10001, new RecordTypes(10001,null) },
				{10002, new RecordTypes(10002,null) },
				// Records ~12000 seem to be related to the Comments used in PPT 2000/XP
				// (Comments in PPT97 are normal Escher text boxes)
				{12000, new RecordTypes(12000, new Comment2000) },
				{12001, new RecordTypes(12001, new Comment2000Atom) },
				{12004, new RecordTypes(12004,null) },
				{12005, new RecordTypes(12005,null) },
				// Records ~12050 seem to be related to Document Encryption
				{12052, new RecordTypes(12052, new DocumentEncryptionAtom) }
			};
		public short typeID;
		public RecordConstructor<Record> RecordConstructor { get; }

		public RecordTypes(int typeID, RecordConstructor<Record> recordConstructor)
		{
			this.typeID = (short)typeID;
			this.RecordConstructor = recordConstructor;
		}

		public static RecordTypes ForTypeID(int typeID)
		{
			RecordTypes rt = RecordTypesLookup[(short)typeID];
			return (rt != null) ? rt : null;
		}
	}

	////@FunctionalInterface
	//public interface RecordConstructor<T>
	//	where T : Record
	//{
	//	T Apply(byte[] source, int start, int len);
	//}

	public delegate T RecordConstructor<T>(byte[] source, int start, int len);
}
