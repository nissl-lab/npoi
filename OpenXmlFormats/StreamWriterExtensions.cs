using System.IO;
using System.Runtime.CompilerServices;

namespace NPOI.OpenXmlFormats;

internal static class StreamWriterExtensions
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteAttribute(this StreamWriter sw, string name, string value)
    {
        sw.Write(" ");
        sw.Write(name);
        sw.Write("=\"");
        sw.Write(value);
        sw.Write("\"");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteAttribute(this StreamWriter sw, string name, int value)
    {
        sw.Write(" ");
        sw.Write(name);
        sw.Write("=\"");
        sw.Write(value);
        sw.Write("\"");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteBooleanAttribute(this StreamWriter sw, string name, bool value)
    {
        sw.Write(" ");
        sw.Write(name);
        sw.Write("=\"");
        sw.Write(value ? 1 : 0);
        sw.Write("\"");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteElementAndContent(this StreamWriter sw, string name, string value)
    {
        sw.Write("<");
        sw.Write(name);
        sw.Write(">");
        sw.Write(value);
        sw.Write("</");
        sw.Write(name);
        sw.Write(">");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteEndElement(this StreamWriter sw, string name)
    {
        sw.Write("</");
        sw.Write(name);
        sw.Write(">");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteEndW(this StreamWriter sw, string nodeName)
    {
        sw.Write("</w:");
        sw.Write(nodeName);
        sw.Write(">");
    }
}