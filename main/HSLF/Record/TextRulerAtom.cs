using NPOI.HSLF.Model;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.HSLF.Record
{
	public class TextRulerAtom : RecordAtom
	{
		private static  BitField DEFAULT_TAB_SIZE = GetInstance(0x0001);
		private static  BitField C_LEVELS = GetInstance(0x0002);
		private static  BitField TAB_STOPS = GetInstance(0x0004);
		private static  BitField[] LEFT_MARGIN_LVL_MASK = {
        GetInstance(0x0008), GetInstance(0x0010), GetInstance(0x0020),
        GetInstance(0x0040), GetInstance(0x0080),
    };
	private static  BitField[] INDENT_LVL_MASK = {
        GetInstance(0x0100), GetInstance(0x0200), GetInstance(0x0400),
        GetInstance(0x0800), GetInstance(0x1000),
    };

/**
 * Record header.
 */
private  byte[] _header = new byte[8];

//ruler internals
private int defaultTabSize;
private int numLevels;
		private List<HSLFTabStop> tabStops = new List<HSLFTabStop>();
//bullet.offset
private  int[] leftMargin = new int[5];
//text.offset
private  int[] indent = new int[5];

/**
 * Constructs a new empty ruler atom.
 */
public TextRulerAtom()
{
	LittleEndian.PutShort(_header, 2, (short)GetRecordType());
}

/**
 * Constructs the ruler atom record from its
 *  source data.
 *
 * @param source the source data as a byte array.
 * @param start the start offset into the byte array.
 * @param len the length of the slice in the byte array.
 */
public TextRulerAtom( byte[] source,  int start,  int len) {
	 LittleEndianByteArrayInputStream leis = new LittleEndianByteArrayInputStream(source, start, Math.Min(len, GetMaxRecordLength()));


	try
	{
		// Get the header.
		IOUtils.ReadFully(leis, _header);

		// Get the record data.
		Read(leis);
	}
	catch (IOException e)
	{
		//LOG.atError().withThrowable(e).log("Failed to parse TextRulerAtom");
	}
}

/**
 * Gets the record type.
 *
 * @return the record type.
 */
//@Override
	public override long GetRecordType()
{
	return RecordTypes.TextRulerAtom.typeID;
}

		/**
		 * Write the contents of the record back, so it can be written
		 * to disk.
		 *
		 * @param out the output stream to write to.
		 * @throws java.io.IOException if an error occurs.
		 */
		//@Override
		public override void WriteOut(OutputStream _out)
		{
			using (MemoryStream bos = new MemoryStream(200))
			{
				LittleEndianOutputStream lbos = new LittleEndianOutputStream(bos);
				int mask = 0;
				mask |= WriteIf(lbos, numLevels, C_LEVELS);
				mask |= WriteIf(lbos, defaultTabSize, DEFAULT_TAB_SIZE);
				mask |= WriteIf(lbos, tabStops, TAB_STOPS);
				for (int i = 0; i < 5; i++)
				{
					mask |= WriteIf(lbos, leftMargin[i], LEFT_MARGIN_LVL_MASK[i]);
					mask |= WriteIf(lbos, indent[i], INDENT_LVL_MASK[i]);
				}
				LittleEndian.PutInt(_header, 4, (int)bos.Length + 4);
				_out.Write(_header);
				LittleEndian.PutUShort(mask, _out);
				LittleEndian.PutUShort(0, _out);
				bos.WriteTo(_out);
			}
		}

		private static int WriteIf(LittleEndianOutputStream lbos, int value, BitField bit)
		{
			bool isSet = false;
			if (value != null || value != 0)
			{
				lbos.WriteShort(value);
				isSet = true;
			}
			return bit.SetBoolean(0, isSet);
		}

		//@SuppressWarnings("SameParameterValue")
		private static int WriteIf(LittleEndianOutputStream lbos, List<HSLFTabStop> value, BitField bit)
		{
			bool isSet = false;
			if (value != null && !(value.Count == 0))
			{
				HSLFTabStopPropCollection.WriteTabStops(lbos, value);
				isSet = true;
			}
			return bit.SetBoolean(0, isSet);
		}

/**
 * Read the record bytes and initialize the internal variables
 */
private void Read( LittleEndianByteArrayInputStream leis)
{
	 int mask = leis.ReadInt();
	numLevels = ReadIf(leis, mask, C_LEVELS);
	defaultTabSize = ReadIf(leis, mask, DEFAULT_TAB_SIZE);
	if (TAB_STOPS.IsSet(mask))
	{
		tabStops.AddRange(HSLFTabStopPropCollection.ReadTabStops(leis));
	}
	for (int i = 0; i < 5; i++)
	{
		leftMargin[i] = ReadIf(leis, mask, LEFT_MARGIN_LVL_MASK[i]);
		indent[i] = ReadIf(leis, mask, INDENT_LVL_MASK[i]);
	}
}

private static int ReadIf( LittleEndianByteArrayInputStream leis,  int mask,  BitField bit)
{
	return (bit.IsSet(mask)) ? (int)leis.ReadShort() : 0;
}

/**
 * Default distance between tab stops, in master coordinates (576 dpi).
 */
public int GetDefaultTabSize()
{
	return defaultTabSize == null ? 0 : defaultTabSize;
}

/**
 * Number of indent levels (maximum 5).
 */
public int GetNumberOfLevels()
{
	return numLevels == null ? 0 : numLevels;
}

/**
 * Default distance between tab stops, in master coordinates (576 dpi).
 */
public List<HSLFTabStop> GetTabStops()
{
	return tabStops;
}

/**
 * Paragraph's distance from shape's left margin, in master coordinates (576 dpi).
 */
public int[] GetTextOffsets()
{
	return leftMargin;
}

/**
 * First line of paragraph's distance from shape's left margin, in master coordinates (576 dpi).
 */
public int[] GetBulletOffsets()
{
	return indent;
}

public static TextRulerAtom GetParagraphInstance()
{
	 TextRulerAtom tra = new TextRulerAtom();
	tra.indent[0] = 249;
	tra.indent[1] = tra.leftMargin[1] = 321;
	return tra;
}

public void SetParagraphIndent(short leftMargin, short indent)
{
	Arrays.Fill<int>(this.leftMargin, 0);
	Arrays.Fill<int>(this.indent, 0);
	this.leftMargin[0] = (int)leftMargin;
	this.indent[0] = (int)indent;
	this.indent[1] = (int)indent;
}

//@Override
	public IDictionary<string, Func<T>> GetGenericProperties<T>()
{
	return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
		"defaultTabSize", () => GetDefaultTabSize(),
		"numLevels", () => GetNumberOfLevels(),
		"tabStops", GetTabStops,
		"leftMargins", ()=>leftMargin,
		"indents", ()=>indent
	);
}
	}
}
