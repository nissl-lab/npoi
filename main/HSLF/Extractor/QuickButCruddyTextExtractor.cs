/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using NPOI.HSSF.Record;
using NPOI.POIFS.FileSystem;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System;
using System.IO;
using NPOI.HSLF.Record;

namespace NPOI.HSLF.Extractor
{
	/**
 * This class will get all the text from a Powerpoint Document, including
 *  all the bits you didn't want, and in a somewhat random order, but will
 *  do it very fast.
 * The class ignores most of the hslf classes, and doesn't use
 *  HSLFSlideShow. Instead, it just does a very basic scan through the
 *  file, grabbing all the text records as it goes. It then returns the
 *  text, either as a single string, or as a vector of all the individual
 *  strings.
 * Because of how it works, it will return a lot of "crud" text that you
 *  probably didn't want! It will return text from master slides. It will
 *  return duplicate text, and some mangled text (powerpoint files often
 *  have duplicate copies of slide text in them). You don't get any idea
 *  what the text was associated with.
 * Almost everyone will want to use @see PowerPointExtractor instead. There
 *  are only a very small number of cases (eg some performance sensitive
 *  lucene indexers) that would ever want to use this!
 */
	public class QuickButCruddyTextExtractor
	{
		private POIFSFileSystem fs;
		private InputStream _is;
		private byte[] pptContents;

		/**
		 * Really basic text extractor, that will also return lots of crud text.
		 * Takes a single argument, the file to extract from
		 */
		//public static void main(String[] args)
		//{
		//	if(args.length< 1) 
		//	{
		//		System.err.println("Usage:");
		//		System.err.println("\tQuickButCruddyTextExtractor <file>");
		//		System.exit(1);
		//	}
			
		//	string file = args[0];
			
		//	QuickButCruddyTextExtractor ppe = new QuickButCruddyTextExtractor(file);
		//	System.out.println(ppe.getTextAsString());
		//	ppe.close();
		//}

		/**
		 * Creates an extractor from a given file name
		 */
		//@SuppressWarnings("resource")
		public QuickButCruddyTextExtractor(string fileName)
			:this(new POIFSFileSystem(new FileInfo(fileName)))
		{
		}

		/**
		 * Creates an extractor from a given input stream
		 */
		//@SuppressWarnings("resource")
		public QuickButCruddyTextExtractor(InputStream iStream)
			:this(new POIFSFileSystem(iStream))
		{
			_is = iStream;
		}

		/**
		 * Creates an extractor from a POIFS Filesystem
		 */
		public QuickButCruddyTextExtractor(POIFSFileSystem poifs) 
		{
			fs = poifs;
			// Find the PowerPoint bit, and get out the bytes
			InputStream pptIs = fs.CreateDocumentInputStream(HSLFSlideShow.POWERPOINT_DOCUMENT);
			pptContents = IOUtils.ToByteArray(pptIs);
			pptIs.Close();
		}


		/**
		 * Shuts down the underlying streams
		 */
		public void Close()
		{
			if(_is != null) { _is.Close(); }
			fs = null;
		}

		/**
		 * Fetches the ALL the text of the powerpoint file, as a single string
		 */
		public string GetTextAsString()
		{
			StringBuilder ret = new StringBuilder();
			List<string> textV = GetTextAsVector();
			foreach (string text in textV) 
			{
				ret.Append(text);
				if (!text.EndsWith("\n"))
				{
					ret.Append('\n');
				}
			}
			return ret.ToString();
		}

		/**
		 * Fetches the ALL the text of the powerpoint file, in a List of
		 *  strings, one per text record
		 */
		public List<String> GetTextAsVector()
		{
			List<String> textV = new List<string>();

			// Set to the start of the file
			int walkPos = 0;

			// Start walking the file, looking for the records
			while (walkPos != -1)
			{
				walkPos = FindTextRecords(walkPos, textV);
			}

			// Return what we find
			return textV;
		}

		/**
		 * For the given position, look if the record is a text record, and wind
		 *  on after.
		 * If it is a text record, grabs out the text. Whatever happens, returns
		 *  the position of the next record, or -1 if no more.
		 */
		public int FindTextRecords(int startPos, List<String> textV)
		{
			// Grab the length, and the first option byte
			// Note that the length doesn't include the 8 byte atom header
			int len = (int)LittleEndian.GetUInt(pptContents, startPos + 4);
			byte opt = pptContents[startPos];

			// If it's a container, step into it and return
			// (If it's a container, option byte 1 BINARY_AND 0x0f will be 0x0f)
			int container = opt & 0x0f;
			if (container == 0x0f)
			{
				return (startPos + 8);
			}

			// Otherwise, check the type to see if it's text
			int type = LittleEndian.GetUShort(pptContents, startPos + 2);

			// TextBytesAtom
			if (type == RecordTypes.TextBytesAtom.typeID)
			{
				TextBytesAtom tba = (TextBytesAtom)Record.Record.CreateRecordForType(type, pptContents, startPos, len + 8);
				string text = HSLFTextParagraph.toExternalString(tba.GetText(), -1);
				textV.Add(text);
			}
			// TextCharsAtom
			if (type == RecordTypes.TextCharsAtom.typeID)
			{
				TextCharsAtom tca = (TextCharsAtom)Record.Record.CreateRecordForType(type, pptContents, startPos, len + 8);
				String text = HSLFTextParagraph.toExternalString(tca.getText(), -1);
				textV.Add(text);
			}

			// CString (doesn't go via a TextRun)
			if (type == RecordTypes.CString.typeID)
			{
				CString cs = (CString)Record.Record.CreateRecordForType(type, pptContents, startPos, len + 8);
				String text = cs.GetText();

				// Ignore the ones we know to be rubbish
				if (!"___PPT10".Equals(text) && !"Default Design".Equals(text))
				{
					textV.Add(text);
				}
			}


			// Wind on by the atom length, and check we're not at the end
			int newPos = (startPos + 8 + len);
			if (newPos > (pptContents.Length - 8))
			{
				newPos = -1;
			}
			return newPos;
		}
	}
}
