using NPOI.SS.Formula.PTG;
using System.Reflection;

namespace testclass
{
    internal class Program
    {
        static void Do()
        {
            const string types = @"ArrayRecord
AutoFilterInfoRecord
BackupRecord
BlankRecord
BOFRecord
BookBoolRecord
BoolErrRecord
BottomMarginRecord
BoundSheetRecord
CalcCountRecord
CalcModeRecord
CFHeaderRecord
CFHeader12Record
CFRuleRecord
CFRule12Record
ChartRecord
AlRunsRecord
CodepageRecord
ColumnInfoRecord
ContinueRecord
CountryRecord
CRNCountRecord
CRNRecord
DateWindow1904Record
DBCellRecord
DConRefRecord
DefaultColWidthRecord
DefaultRowHeightRecord
DeltaRecord
DimensionsRecord
DrawingGroupRecord
DrawingRecord
DrawingSelectionRecord
DSFRecord
DVALRecord
DVRecord
EOFRecord
ExtendedFormatRecord
ExternalNameRecord
ExternSheetRecord
ExtSSTRecord
FilePassRecord
FileSharingRecord
FnGroupCountRecord
FontRecord
FooterRecord
FormatRecord
FormulaRecord
GridsetRecord
GutsRecord
HCenterRecord
HeaderRecord
HeaderFooterRecord
HideObjRecord
HorizontalPageBreakRecord
HyperlinkRecord
IndexRecord
InterfaceEndRecord
InterfaceHdrRecord
IterationRecord
LabelRecord
LabelSSTRecord
LeftMarginRecord
MergeCellsRecord
MMSRecord
MulBlankRecord
MulRKRecord
NameRecord
NameCommentRecord
NoteRecord
NumberRecord
ObjectProtectRecord
ObjRecord
PaletteRecord
PaneRecord
PasswordRecord
PasswordRev4Record
PrecisionRecord
PrintGridlinesRecord
PrintHeadersRecord
PrintSetupRecord
PrintSizeRecord
ProtectionRev4Record
ProtectRecord
RecalcIdRecord
RefModeRecord
RefreshAllRecord
RightMarginRecord
RKRecord
RowRecord
SaveRecalcRecord
ScenarioProtectRecord
SCLRecord
SelectionRecord
SeriesRecord
SeriesTextRecord
SharedFormulaRecord
SSTRecord
StringRecord
StyleRecord
SupBookRecord
TabIdRecord
TableRecord
TableStylesRecord
TextObjectRecord
TopMarginRecord
UncalcedRecord
UseSelFSRecord
UserSViewBegin
UserSViewEnd
ValueRangeRecord
VCenterRecord
VerticalPageBreakRecord
WindowOneRecord
WindowProtectRecord
WindowTwoRecord
WriteAccessRecord
WriteProtectRecord
WSBoolRecord
SheetExtRecord
AreaFormatRecord
AreaRecord
AttachedLabelRecord
AxcExtRecord
AxisLineFormatRecord
AxisParentRecord
AxisRecord
AxesUsedRecord
BarRecord
BeginRecord
BopPopCustomRecord
BopPopRecord
CatLabRecord
CatSerRangeRecord
Chart3DBarShapeRecord
Chart3dRecord
ChartEndObjectRecord
ChartFormatRecord
ChartFRTInfoRecord
ChartStartObjectRecord
CrtLayout12ARecord
CrtLayout12Record
CrtLineRecord
CrtLinkRecord
CrtMlFrtContinueRecord
CrtMlFrtRecord
DataFormatRecord
DataLabExtContentsRecord
DataLabExtRecord
DatRecord
DefaultTextRecord
DropBarRecord
EndBlockRecord
EndRecord
LinkedDataRecord
Fbi2Record
FbiRecord
FontIndexRecord
FrameRecord
FrtFontListRecord
GelFrameRecord
IFmtRecordRecord
LegendExceptionRecord
LegendRecord
LineFormatRecord
MarkerFormatRecord
ObjectLinkRecord
PicFRecord
PieFormatRecord
PieRecord
PlotAreaRecord
PlotGrowthRecord
PosRecord
RadarAreaRecord
RadarRecord
RichTextStreamRecord
ScatterRecord
SerAuxErrBarRecord
SerAuxTrendRecord
SerFmtRecord
SeriesIndexRecord
SeriesListRecord
SerParentRecord
SerToCrtRecord
ShapePropsStreamRecord
ShtPropsRecord
StartBlockRecord
SurfRecord
TextPropsStreamRecord
TextRecord
TickRecord
UnitsRecord
YMultRecord
DataItemRecord
ExtendedPivotTableViewFieldsRecord
PageItemRecord
StreamIDRecord
ViewDefinitionRecord
ViewFieldsRecord
ViewSourceRecord
AutoFilterRecord
FilterModeRecord
Excel9FileRecord";

            foreach(var it in types.Split(new[] {"\r", "\n"}, StringSplitOptions.RemoveEmptyEntries))
            {
                Console.WriteLine($"_recordTypes.Add({it}.sid, typeof({it}))");
            }
        }
        static void Main(string[] args)
        {
            Do();
            return;
            var parentType = typeof(OperandPtg);

            var types = AppDomain.CurrentDomain.GetAssemblies().SelectMany(a => a.GetTypes()).Where(a => parentType.IsAssignableFrom(a));

            foreach (var type in types)
            {
                Console.WriteLine(type.FullName);
                Console.WriteLine("======================");

                foreach (var field in type.GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
                {
                    var fieldType = field.FieldType;

                    if (!IsValueType(fieldType))
                    {
                        Console.WriteLine($"{field.Name}: {fieldType.FullName}");
                    }
                }

                Console.WriteLine();
                Console.WriteLine();
            }
        }

        static bool IsValueType(Type type)
        {
            if (type == typeof(string)) return true;
            return type.IsValueType;
        }
    }
}