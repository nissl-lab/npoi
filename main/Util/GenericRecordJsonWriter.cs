using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using MathNet.Numerics.LinearAlgebra;
using NPOI.Common.UserModel;
using NPOI.HSLF.Record;
using NPOI.SL.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.Util.Collections;
using SixLabors.ImageSharp.Memory;
using static NPOI.HSSF.Util.HSSFColor;

namespace NPOI.Util
{
	public class GenericRecordJsonWriter : ICloseable
	{
		static char[] t = new char[255];
		private static string TABS;
		private static string ZEROS = "0000000000000000";
		private static Regex ESC_CHARS = new Regex(@"[\\\p{C}\\\\]");
		private static string NL = Environment.NewLine;


		//Arrays.Fill(t, '\t')

		/**
         * Handler method
         *
         * @param record the parent record, applied via instance method reference
         * @param name the name of the property
         * @param object the value of the property
         * @return {@code true}, if the element was handled and output produced,
         *   The provided methods can be overridden and a implementation can return {@code false},
         *   if the element hasn't been written to the stream
         */
		protected delegate bool GenericRecordHandler(GenericRecordJsonWriter record, string name, object o);


		protected AppendableWriter aw;
		protected TextWriter fw;
		protected int indent = 0;
		protected bool withComments = true;
		protected int childIndex = 0;

		public GenericRecordJsonWriter(FileInfo fileName)
		{
			OutputStream os = ("null".Equals(fileName.Name)) ? null : File.OpenWrite(fileName.FullName);
			aw = new AppendableWriter(new OutputStreamWriter(os, StandardCharsets.UTF_8));
			fw = new PrintWriter(aw);
		}

		public GenericRecordJsonWriter(Appendable buffer)
		{
			aw = new AppendableWriter(buffer);
			fw = new PrintWriter(aw);
		}

		public static String marshal(GenericRecord record)
		{
			return marshal(record, true);
		}

		public static String marshal(GenericRecord record, bool withComments)
		{
			StringBuilder sb = new StringBuilder();
			try
			{

				using (GenericRecordJsonWriter w = new GenericRecordJsonWriter(sb))
				{
					w.SetWithComments(withComments);
					w.Write(record);
					return sb.toString();
				}
			}
			catch (IOException e)
			{
				return "{}";
			}
		}

		public void SetWithComments(bool withComments)
		{
			this.withComments = withComments;
		}

		public void Close()
		{
			fw.close();
		}

		protected String tabs()
		{
			return TABS.substring(0, Math.min(indent, TABS.length()));
		}

		public void Write(GenericRecord record)
		{
			string tabs = tabs();
			RecordTypes type = record.GetGenericRecordType();
			String recordName = (type != null) ? type.Name() : record.GetClass().getSimpleName();
			fw.append(tabs);
			fw.append("{");
			if (withComments)
			{
				fw.append("   /* ");
				fw.append(recordName);
				if (childIndex > 0)
				{
					fw.append(" - index: ");
					fw.print(childIndex);
				}
				fw.append(" */");
			}
			fw.println();

			bool hasProperties = writeProperties(record);
			fw.println();

			writeChildren(record, hasProperties);

			fw.append(tabs);
			fw.append("}");
		}

		protected bool writeProperties(GenericRecord record)
		{
			Map < String, Supplier <?>> prop = record.getGenericProperties();
			if (prop == null || prop.isEmpty())
			{
				return false;
			}

			final int oldChildIndex = childIndex;
			childIndex = 0;
			long cnt = prop.entrySet().stream().filter(e->writeProp(e.getKey(), e.getValue())).count();
			childIndex = oldChildIndex;

			return cnt > 0;
		}


		protected bool writeChildren(GenericRecord record, bool hasProperties)
		{
			List <? extends GenericRecord > list = record.getGenericChildren();
			if (list == null || list.isEmpty())
			{
				return false;
			}

			indent++;
			aw.setHoldBack(tabs() + (hasProperties ? ", " : "") + "\"children\": [" + NL);
			final int oldChildIndex = childIndex;
			childIndex = 0;
			long cnt = list.stream().filter(l->writeValue(null, l) && ++childIndex > 0).count();
			childIndex = oldChildIndex;
			aw.setHoldBack(null);

			if (cnt > 0)
			{
				fw.println();
				fw.println(tabs() + "]");
			}
			indent--;

			return cnt > 0;
		}

		public void writeError(String errorMsg)
		{
			fw.append("{ error: ");
			printObject("error", errorMsg);
			fw.append(" }");
		}

		protected bool writeProp(String name, Supplier<?> value)
		{
			final bool isNext = (childIndex > 0);
			aw.setHoldBack(isNext ? NL + tabs() + "\t, " : tabs() + "\t  ");
			final int oldChildIndex = childIndex;
			childIndex = 0;
			bool written = writeValue(name, value.get());
			childIndex = oldChildIndex + (written ? 1 : 0);
			aw.setHoldBack(null);
			return written;
		}

		protected bool writeValue(String name, Object o)
		{
			if (childIndex > 0)
			{
				aw.setHoldBack(",");
			}

			GenericRecordHandler grh = (o == null)
			? GenericRecordJsonWriter::printNull
				: handler.stream().filter(h->matchInstanceOrArray(h.getKey(), o)).
				findFirst().map(Map.Entry::getValue).orElse(null);

			bool result = grh != null && grh.print(this, name, o);
			aw.setHoldBack(null);
			return result;
		}

		protected static bool matchInstanceOrArray(Class<?> key, Object instance)
		{
			return key.isInstance(instance) || (Array.class.equals(key) && instance.getClass().isArray());
    }

	protected void printName(String name)
	{
		fw.print(name != null ? "\"" + name + "\": " : "");
	}

	protected bool printNull(String name, Object o)
	{
		printName(name);
		fw.write("null");
		return true;
	}

	//@SuppressWarnings("java:S3516")
	protected bool printNumber(String name, Object o)
	{
		Number n = (Number)o;
		printName(name);

		if (o instanceof Float) {
			fw.print(n.floatValue());
			return true;
		} else if (o instanceof Double) {
			fw.print(n.doubleValue());
			return true;
		}

		fw.print(n.longValue());

		final int size;
		if (n instanceof Byte) {
			size = 2;
		} else if (n instanceof Short) {
			size = 4;
		} else if (n instanceof Integer) {
			size = 8;
		} else if (n instanceof Long) {
			size = 16;
		} else
		{
			size = -1;
		}

		long l = n.longValue();
		if (withComments && size > 0 && (l < 0 || l > 9))
		{
			fw.write(" /* 0x");
			fw.write(trimHex(l, size));
			fw.write(" */");
		}
		return true;
	}

	protected bool printbool(String name, Object o)
	{
		printName(name);
		fw.write(((bool)o).toString());
		return true;
	}

	protected bool printList(String name, Object o)
	{
		printName(name);
		fw.println("[");
		int oldChildIndex = childIndex;
		childIndex = 0;
		((List <?>)o).forEach(e-> { writeValue(null, e); childIndex++; });
		childIndex = oldChildIndex;
		fw.write(tabs() + "\t]");
		return true;
	}

	protected bool printGenericRecord(String name, Object o)
	{
		printName(name);
		this.indent++;
		write((GenericRecord)o);
		this.indent--;
		return true;
	}

	protected bool printAnnotatedFlag(String name, Object o)
	{
		printName(name);
		AnnotatedFlag af = (AnnotatedFlag)o;
		fw.print(af.getValue().get().longValue());
		if (withComments)
		{
			fw.write(" /* ");
			fw.write(af.getDescription());
			fw.write(" */ ");
		}
		return true;
	}

	protected bool printBytes(String name, Object o)
	{
		printName(name);
		fw.write('"');
		fw.write(Base64.getEncoder().encodeToString((byte[])o));
		fw.write('"');
		return true;
	}

	protected bool printPoint(String name, Object o)
	{
		printName(name);
		Point2D p = (Point2D)o;
		fw.write("{ \"x\": " + p.getX() + ", \"y\": " + p.getY() + " }");
		return true;
	}

	protected bool printDimension(String name, Object o)
	{
		printName(name);
		Dimension2D p = (Dimension2D)o;
		fw.write("{ \"width\": " + p.getWidth() + ", \"height\": " + p.getHeight() + " }");
		return true;
	}

	protected bool printRectangle(String name, Object o)
	{
		printName(name);
		Rectangle2D p = (Rectangle2D)o;
		fw.write("{ \"x\": " + p.getX() + ", \"y\": " + p.getY() + ", \"width\": " + p.getWidth() + ", \"height\": " + p.getHeight() + " }");
		return true;
	}

	protected bool printPath(String name, Object o)
	{
		printName(name);
		final PathIterator iter = ((Path2D)o).getPathIterator(null);
		final double[] pnts = new double[6];
		fw.write("[");

		indent += 2;
		String t = tabs();
		indent -= 2;

		bool isNext = false;
		while (!iter.isDone())
		{
			fw.println(isNext ? ", " : "");
			fw.print(t);
			isNext = true;
			final int segType = iter.currentSegment(pnts);
			fw.append("{ \"type\": ");
			switch (segType)
			{
				case PathIterator.SEG_MOVETO:
					fw.write("\"move\", \"x\": " + pnts[0] + ", \"y\": " + pnts[1]);
					break;
				case PathIterator.SEG_LINETO:
					fw.write("\"lineto\", \"x\": " + pnts[0] + ", \"y\": " + pnts[1]);
					break;
				case PathIterator.SEG_QUADTO:
					fw.write("\"quad\", \"x1\": " + pnts[0] + ", \"y1\": " + pnts[1] + ", \"x2\": " + pnts[2] + ", \"y2\": " + pnts[3]);
					break;
				case PathIterator.SEG_CUBICTO:
					fw.write("\"cubic\", \"x1\": " + pnts[0] + ", \"y1\": " + pnts[1] + ", \"x2\": " + pnts[2] + ", \"y2\": " + pnts[3] + ", \"x3\": " + pnts[4] + ", \"y3\": " + pnts[5]);
					break;
				case PathIterator.SEG_CLOSE:
					fw.write("\"close\"");
					break;
			}
			fw.append(" }");
			iter.next();
		}

		fw.write("]");
		return true;
	}

	protected bool printObject(String name, Object o)
	{
		printName(name);
		fw.write('"');

		final String str = o.toString();
		final Matcher m = ESC_CHARS.matcher(str);
		int pos = 0;
		while (m.find())
		{
			fw.append(str, pos, m.start());
			String match = m.group();
			switch (match)
			{
				case "\n":
					fw.write("\\\\n");
					break;
				case "\r":
					fw.write("\\\\r");
					break;
				case "\t":
					fw.write("\\\\t");
					break;
				case "\b":
					fw.write("\\\\b");
					break;
				case "\f":
					fw.write("\\\\f");
					break;
				case "\\":
					fw.write("\\\\\\\\");
					break;
				case "\"":
					fw.write("\\\\\"");
					break;
				default:
					fw.write("\\\\u");
					fw.write(trimHex(match.charAt(0), 4));
					break;
			}
			pos = m.end();
		}
		fw.append(str, pos, str.length());
		fw.write('"');
		return true;
	}

	protected bool printAffineTransform(String name, Object o)
	{
		printName(name);
		AffineTransform xForm = (AffineTransform)o;
		fw.write(
			"{ \"scaleX\": " + xForm.getScaleX() +
			", \"shearX\": " + xForm.getShearX() +
			", \"transX\": " + xForm.getTranslateX() +
			", \"scaleY\": " + xForm.getScaleY() +
			", \"shearY\": " + xForm.getShearY() +
			", \"transY\": " + xForm.getTranslateY() + " }");
		return true;
	}

	protected bool printColor(String name, Object o)
	{
		printName(name);

		final int rgb = ((Color)o).getRGB();
		fw.print(rgb);

		if (withComments)
		{
			fw.write(" /* 0x");
			fw.write(trimHex(rgb, 8));
			fw.write(" */");
		}
		return true;
	}

	protected bool printArray(String name, Object o)
	{
		printName(name);
		fw.write("[");
		int length = Array.getLength(o);
		final int oldChildIndex = childIndex;
		for (childIndex = 0; childIndex < length; childIndex++)
		{
			writeValue(null, Array.get(o, childIndex));
		}
		childIndex = oldChildIndex;
		fw.write(tabs() + "\t]");
		return true;
	}

	protected bool printImage(String name, Object o)
	{
		BufferedImage img = (BufferedImage)o;

		final String[] COLOR_SPACES = {
			"XYZ","Lab","Luv","YCbCr","Yxy","RGB","GRAY","HSV","HLS","CMYK","Unknown","CMY","Unknown"

		};

		final String[] IMAGE_TYPES = {
			"CUSTOM","INT_RGB","INT_ARGB","INT_ARGB_PRE","INT_BGR","3BYTE_BGR","4BYTE_ABGR","4BYTE_ABGR_PRE",
                "USHORT_565_RGB","USHORT_555_RGB","BYTE_GRAY","USHORT_GRAY","BYTE_BINARY","BYTE_INDEXED"

		};

		printName(name);
		ColorModel cm = img.getColorModel();
		String colorType =
			(cm instanceof IndexColorModel) ? "indexed" :
            (cm instanceof ComponentColorModel) ? "component" :
            (cm instanceof DirectColorModel) ? "direct" :
            (cm instanceof PackedColorModel) ? "packed" : "unknown";
		fw.write(
			"{ \"width\": " + img.getWidth() +
			", \"height\": " + img.getHeight() +
			", \"type\": \"" + IMAGE_TYPES[img.getType()] + "\"" +
			", \"colormodel\": \"" + colorType + "\"" +
			", \"pixelBits\": " + cm.getPixelSize() +
			", \"numComponents\": " + cm.getNumComponents() +
			", \"colorSpace\": \"" + COLOR_SPACES[Math.min(cm.getColorSpace().getType(), 12)] + "\"" +
			", \"transparency\": " + cm.getTransparency() +
			", \"alpha\": " + cm.hasAlpha() +
			"}"
		);
		return true;
	}

	static String trimHex(final long l, final int size)
	{
		final String b = Long.toHexString(l);
		int len = b.length();
		return ZEROS.substring(0, Math.max(0, size - len)) + b.substring(Math.max(0, len - size), len);
	}
}

	class AppendableWriter : TextWriter
	{

		//private Appendable appender;
		private TextWriter writer;
		private String holdBack;

	public override Encoding Encoding => throw new NotImplementedException();

	AppendableWriter(Appendable buffer) {
	super(buffer);
	this.appender = buffer;
	this.writer = null;
}

AppendableWriter(TextWriter writer) {
	super(writer);
	this.appender = null;
	this.writer = writer;
}

void setHoldBack(String holdBack)
{
	this.holdBack = holdBack;
}

//@Override
		public void write(char[] cbuf, int off, int len)
            if (holdBack != null) {
		if (appender != null)
		{
			appender.append(holdBack);
		}
		else if (writer != null)
		{
			writer.write(holdBack);
		}
		holdBack = null;
	}

            if (appender != null) {
		appender.append(String.valueOf(cbuf), off, len);
	} else if (writer != null) {
		writer.write(cbuf, off, len);
	}
}

//@Override
		public void flush() throws IOException
{
	Object o = (appender != null) ? appender : writer;
            if (o instanceof Flushable) {
		((Flushable)o).flush();
	}
}

//@Override
		public void close() throws IOException
{
	flush();
	Object o = (appender != null) ? appender : writer;
            if (o instanceof Closeable) {
		((Closeable)o).close();
	}
}
    }

}
