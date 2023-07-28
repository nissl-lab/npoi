using System;
using System.Text;

namespace testclass
{
    internal class Program
    {
        static void Do()
        {
            const string types = @"XWPFDocument
XSSFWorkbook
XWPFComments
XWPFFootnotes
XWPFFooter
XWPFHeader
XWPFNumbering
XWPFPictureData
XWPFSettings
XWPFStyles
XSSFChart
XSSFDrawing
XSSFPictureData
XSSFPivotCache
XSSFPivotCacheDefinition
XSSFPivotCacheRecords
XSSFPivotTable
XSSFSheet
XSSFChartSheet
XSSFDialogsheet
XSSFTable
XSSFVBAPart
XSSFVMLDrawing
CalculationChain
CommentsTable
ExternalLinksTable
MapInfo
SharedStringsTable
SingleXmlCells
StylesTable
ThemesTable";

            var cls = types.Split(new[] { "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            var sb1 = new StringBuilder();
            var sb2 = new StringBuilder();
            var sb3 = new StringBuilder();

            foreach (var clsname in cls)
            {
                sb1.AppendLine($"if (cls == typeof({clsname}))");
                sb2.AppendLine($"if (cls == typeof({clsname}))");
                sb3.AppendLine($"if (cls == typeof({clsname}))");

                sb1.AppendLine($"return new {clsname}();");
                sb2.AppendLine($"return new {clsname}(part);");
                sb3.AppendLine($"return new {clsname}(parent, part);");
            }

            Console.WriteLine("Func1:");
            Console.WriteLine(sb1.ToString());
            Console.WriteLine();
            Console.WriteLine("Func2:");
            Console.WriteLine(sb2.ToString());
            Console.WriteLine();
            Console.WriteLine("Func3:");
            Console.WriteLine(sb3.ToString());

            /*foreach (var it in )
            {
                Console.WriteLine($"_recordTypes.Add({it}.sid, typeof({it}))");
            }*/
        }
        static void Main(string[] args)
        {
            Do();
            return;
            /*var parentType = typeof(OperandPtg);

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
            }*/
        }

        static bool IsValueType(Type type)
        {
            if (type == typeof(string)) return true;
            return type.IsValueType;
        }
    }
}