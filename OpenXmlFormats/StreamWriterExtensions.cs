using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

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
    public static async Task WriteAttributeAsync(this StreamWriter sw, string name, string value)
    {
        await sw.WriteAsync(" ");
        await sw.WriteAsync(name);
        await sw.WriteAsync("=\"");
        await sw.WriteAsync(value);
        await sw.WriteAsync("\"");
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
    public static async Task WriteAttributeAsync(this StreamWriter sw, string name, int value)
    {
        await sw.WriteAsync(" ");
        await sw.WriteAsync(name);
        await sw.WriteAsync("=\"");
        await sw.WriteAsync(value.ToString());
        await sw.WriteAsync("\"");
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
    public static async Task WriteBooleanAttributeAsync(this StreamWriter sw, string name, bool value)
    {
        await sw.WriteAsync(" ");
        await sw.WriteAsync(name);
        await sw.WriteAsync("=\"");
        await sw.WriteAsync(value ? '1':'0');
        await sw.WriteAsync("\"");
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
    public static async Task WriteElementAndContentAsync(this StreamWriter sw, string name, string value)
    {
        await sw.WriteAsync("<");
        await sw.WriteAsync(name);
        await sw.WriteAsync(">");
        await sw.WriteAsync(value);
        await sw.WriteAsync("</");
        await sw.WriteAsync(name);
        await sw.WriteAsync(">");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteEndElement(this StreamWriter sw, string name)
    {
        sw.Write("</");
        sw.Write(name);
        sw.Write(">");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static async Task WriteEndElementAsync(this StreamWriter sw, string name)
    {
        await sw.WriteAsync("</");
        await sw.WriteAsync(name);
        await sw.WriteAsync(">");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteEndW(this StreamWriter sw, string nodeName)
    {
        sw.Write("</w:");
        sw.Write(nodeName);
        sw.Write(">");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static async Task WriteEndWAsync(this StreamWriter sw, string nodeName)
    {
        await sw.WriteAsync("</w:");
        await sw.WriteAsync(nodeName);
        await sw.WriteAsync(">");
    }
}