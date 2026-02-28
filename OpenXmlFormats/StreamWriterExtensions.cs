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
        sw.Write(value ? '1' : '0');
        sw.Write("\"");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteElementAndContent(this StreamWriter sw, string name, string value)
    {
        sw.Write("<");
        sw.Write(name);
        sw.Write('>');
        sw.Write(value);
        sw.Write("</");
        sw.Write(name);
        sw.Write('>');
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteStart(this StreamWriter sw, string nodeName)
    {
        sw.Write('<');
        sw.Write(nodeName);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteStart(this StreamWriter sw, string ns, string nodeName)
    {
        sw.Write('<');
        sw.Write(ns);
        sw.Write(':');
        sw.Write(nodeName);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteEndElement(this StreamWriter sw, string name)
    {
        sw.Write("</");
        sw.Write(name);
        sw.Write('>');
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteEndElement(this StreamWriter sw, string ns, string nodeName)
    {
        sw.Write("</");
        sw.Write(ns);
        sw.Write(':');
        sw.Write(nodeName);
        sw.Write('>');
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteStartW(this StreamWriter sw, string nodeName)
    {
        sw.WriteStart("w", nodeName);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteEndW(this StreamWriter sw, string nodeName)
    {
        sw.WriteEndElement("w", nodeName);
    }
}
