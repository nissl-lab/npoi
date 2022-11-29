using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.Util
{
    public class ClassIDPredefined
	{
		public static readonly Dictionary<string, (string externalForm, string fileExtension, string contentType)> ClassIDPredefinedLookup =
			new Dictionary<string, (string externalForm, string fileExtension, string contentType)>
			{
				/** OLE 1.0 package manager */
				{ "OLE_V1_PACKAGE",("{0003000C-0000-0000-C000-000000000046}", ".bin", null) },
				/** Excel V3 - document */
				{ "EXCEL_V3",("{00030000-0000-0000-C000-000000000046}", ".xls", "application/vnd.ms-excel") },
				/** Excel V3 - chart */
				{ "EXCEL_V3_CHART",("{00030001-0000-0000-C000-000000000046}", null, null) },
				/** Excel V3 - macro */
				{ "EXCEL_V3_MACRO",("{00030002-0000-0000-C000-000000000046}", null, null) },
				/** Excel V7 / 95 - document */
				{ "EXCEL_V7",("{00020810-0000-0000-C000-000000000046}", ".xls", "application/vnd.ms-excel") },
				/** Excel V7 / 95 - workbook */
				{ "EXCEL_V7_WORKBOOK",("{00020841-0000-0000-C000-000000000046}", null, null) },
				/** Excel V7 / 95 - chart */
				{ "EXCEL_V7_CHART",("{00020811-0000-0000-C000-000000000046}", null, null) },
				/** Excel V8 / 97 - document */
				{ "EXCEL_V8",("{00020820-0000-0000-C000-000000000046}", ".xls", "application/vnd.ms-excel") },
				/** Excel V8 / 97 - chart */
				{ "EXCEL_V8_CHART",("{00020821-0000-0000-C000-000000000046}", null, null) },
				/** Excel V11 / 2003 - document */
				{ "EXCEL_V11",("{00020812-0000-0000-C000-000000000046}", ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") },
				/** Excel V12 / 2007 - document */
				{ "EXCEL_V12",("{00020830-0000-0000-C000-000000000046}", ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") },
				/** Excel V12 / 2007 - macro */
				{ "EXCEL_V12_MACRO",("{00020832-0000-0000-C000-000000000046}", ".xlsm", "application/vnd.ms-excel.sheet.macroEnabled.12") },
				/** Excel V12 / 2007 - xlsb document */
				{ "EXCEL_V12_XLSB",("{00020833-0000-0000-C000-000000000046}", ".xlsb", "application/vnd.ms-excel.sheet.binary.macroEnabled.12") },
				/* Excel V14 / 2010 - document */
				{ "EXCEL_V14",("{00024500-0000-0000-C000-000000000046}", ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") },
				/* Excel V14 / 2010 - workbook */
				{ "EXCEL_V14_WORKBOOK",("{000208D5-0000-0000-C000-000000000046}", null, null) },
				/* Excel V14 / 2010 - chart */
				{ "EXCEL_V14_CHART",("{00024505-0014-0000-C000-000000000046}", null, null) },
				/** Excel V14 / 2010 - OpenDocument spreadsheet */
				{ "EXCEL_V14_ODS",("{EABCECDB-CC1C-4A6F-B4E3-7F888A5ADFC8}", ".ods", "application/vnd.oasis.opendocument.spreadsheet") },
				/** Word V7 / 95 - document */
				{ "WORD_V7",("{00020900-0000-0000-C000-000000000046}", ".doc", "application/msword") },
				/** Word V8 / 97 - document */
				{ "WORD_V8",("{00020906-0000-0000-C000-000000000046}", ".doc", "application/msword") },
				/** Word V12 / 2007 - document */
				{ "WORD_V12",("{F4754C9B-64F5-4B40-8AF4-679732AC0607}", ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document") },
				/** Word V12 / 2007 - macro */
				{ "WORD_V12_MACRO",("{18A06B6B-2F3F-4E2B-A611-52BE631B2D22}", ".docm", "application/vnd.ms-word.document.macroEnabled.12") },
				/** Powerpoint V7 / 95 - document */
				{ "POWERPOINT_V7",("{EA7BAE70-FB3B-11CD-A903-00AA00510EA3}", ".ppt", "application/vnd.ms-powerpoint") },
				/** Powerpoint V7 / 95 - slide */
				{ "POWERPOINT_V7_SLIDE",("{EA7BAE71-FB3B-11CD-A903-00AA00510EA3}", null, null) },
				/** Powerpoint V8 / 97 - document */
				{ "POWERPOINT_V8",("{64818D10-4F9B-11CF-86EA-00AA00B929E8}", ".ppt", "application/vnd.ms-powerpoint") },
				/** Powerpoint V8 / 97 - template */
				{ "POWERPOINT_V8_TPL",("{64818D11-4F9B-11CF-86EA-00AA00B929E8}", ".pot", "application/vnd.ms-powerpoint") },
				/** Powerpoint V12 / 2007 - document */
				{ "POWERPOINT_V12",("{CF4F55F4-8F87-4D47-80BB-5808164BB3F8}", ".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation") },
				/** Powerpoint V12 / 2007 - macro */
				{ "POWERPOINT_V12_MACRO",("{DC020317-E6E2-4A62-B9FA-B3EFE16626F4}", ".pptm", "application/vnd.ms-powerpoint.presentation.macroEnabled.12") },
				/** Publisher V12 */
				{ "PUBLISHER_V12",("{0002123D-0000-0000-C000-000000000046}", ".pub", "application/x-mspublisher") },
				/** Visio 2000 (V6) / 2002 (V10) - drawing */
				{ "VISIO_V10",("{00021A14-0000-0000-C000-000000000046}", ".vsd", "application/vnd.visio") },
				/** Equation Editor 3.0 */
				{ "EQUATION_V3",("{0002CE02-0000-0000-C000-000000000046}", null, null) },
				/** AcroExch.Document */
				{ "PDF",("{B801CA65-A1FC-11D0-85AD-444553540000}", ".pdf", "application/pdf") },
				/** Plain Text Persistent Handler **/
				{ "TXT_ONLY",("{5e941d80-bf96-11cd-b579-08002b30bfeb}", ".txt", "text/plain") },
				/** Microsoft Paint **/
				{ "PAINT",("{0003000A-0000-0000-C000-000000000046}", null, null) },
				/** Standard Hyperlink / STD Moniker **/
				{ "STD_MONIKER",("{79EAC9D0-BAF9-11CE-8C82-00AA004BA90B}", null, null) },
				/** URL Moniker **/
				{ "URL_MONIKER",("{79EAC9E0-BAF9-11CE-8C82-00AA004BA90B}", null, null) },
				/** File Moniker **/
				{ "FILE_MONIKER",("{00000303-0000-0000-C000-000000000046}", null, null) },
				/** Document summary information first property section **/
				{ "DOC_SUMMARY",("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", null, null) },
				/** Document summary information user defined properties section **/
				{ "USER_PROPERTIES",("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", null, null) },
				/** Summary information property section **/
				{ "SUMMARY_PROPERTIES",("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", null, null) }
			};

		//private static Dictionary<string, ClassIDPredefined> LOOKUP = new Dictionary<string, ClassIDPredefined>();
		public static ClassIDPredefined OLE_V1_PACKAGE = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "OLE_V1_PACKAGE").Value);
		/** Excel V3 - document */
		public static ClassIDPredefined EXCEL_V3 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V3").Value);
		/** Excel V3 - chart */
		public static ClassIDPredefined EXCEL_V3_CHART = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V3_CHART").Value);
		/** Excel V3 - macro */
		public static ClassIDPredefined EXCEL_V3_MACRO = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V3_MACRO").Value);
		/** Excel V7 / 95 - document */
		public static ClassIDPredefined EXCEL_V7 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V7").Value);
		/** Excel V7 / 95 - workbook */
		public static ClassIDPredefined EXCEL_V7_WORKBOOK = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V7_WORKBOOK").Value);
		/** Excel V7 / 95 - chart */
		public static ClassIDPredefined EXCEL_V7_CHART = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V7_CHART").Value);
		/** Excel V8 / 97 - document */
		public static ClassIDPredefined EXCEL_V8 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V8").Value);
		/** Excel V8 / 97 - chart */
		public static ClassIDPredefined EXCEL_V8_CHART = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V8_CHART").Value);
		/** Excel V11 / 2003 - document */
		public static ClassIDPredefined EXCEL_V11 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V11").Value);
		/** Excel V12 / 2007 - document */
		public static ClassIDPredefined EXCEL_V12 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V12").Value);
		/** Excel V12 / 2007 - macro */
		public static ClassIDPredefined EXCEL_V12_MACRO = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V12_MACRO").Value);
		/** Excel V12 / 2007 - xlsb document */
		public static ClassIDPredefined EXCEL_V12_XLSB = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V12_XLSB").Value);
		/* Excel V14 / 2010 - document */
		public static ClassIDPredefined EXCEL_V14 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V14").Value);
		/* Excel V14 / 2010 - workbook */
		public static ClassIDPredefined EXCEL_V14_WORKBOOK = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V14_WORKBOOK").Value);
		/* Excel V14 / 2010 - chart */
		public static ClassIDPredefined EXCEL_V14_CHART = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V14_CHART").Value);
		/** Excel V14 / 2010 - OpenDocument spreadsheet */
		public static ClassIDPredefined EXCEL_V14_ODS = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EXCEL_V14_ODS").Value);
		/** Word V7 / 95 - document */
		public static ClassIDPredefined WORD_V7 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "WORD_V7").Value);
		/** Word V8 / 97 - document */
		public static ClassIDPredefined WORD_V8 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "WORD_V8").Value);
		/** Word V12 / 2007 - document */
		public static ClassIDPredefined WORD_V12 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "WORD_V12").Value);
		/** Word V12 / 2007 - macro */
		public static ClassIDPredefined WORD_V12_MACRO = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "WORD_V12_MACRO").Value);
		/** Powerpoint V7 / 95 - document */
		public static ClassIDPredefined POWERPOINT_V7 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "POWERPOINT_V7").Value);
		/** Powerpoint V7 / 95 - slide */
		public static ClassIDPredefined POWERPOINT_V7_SLIDE = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "POWERPOINT_V7_SLIDE").Value);
		/** Powerpoint V8 / 97 - document */
		public static ClassIDPredefined POWERPOINT_V8 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "POWERPOINT_V8").Value);
		/** Powerpoint V8 / 97 - template */
		public static ClassIDPredefined POWERPOINT_V8_TPL = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "POWERPOINT_V8_TPL").Value);
		/** Powerpoint V12 / 2007 - document */
		public static ClassIDPredefined POWERPOINT_V12 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "POWERPOINT_V12").Value);
		/** Powerpoint V12 / 2007 - macro */
		public static ClassIDPredefined POWERPOINT_V12_MACRO = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "POWERPOINT_V12_MACRO").Value);
		/** Publisher V12 */
		public static ClassIDPredefined PUBLISHER_V12 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "PUBLISHER_V12").Value);
		/** Visio 2000 (V6) / 2002 (V10) - drawing */
		public static ClassIDPredefined VISIO_V10 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "VISIO_V10").Value);
		/** Equation Editor 3.0 */
		public static ClassIDPredefined EQUATION_V3 = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "EQUATION_V3").Value);
		/** AcroExch.Document */
		public static ClassIDPredefined PDF = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "PDF").Value);
		/** Plain Text Persistent Handler **/
		public static ClassIDPredefined TXT_ONLY = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "TXT_ONLY").Value);
		/** Microsoft Paint **/
		public static ClassIDPredefined PAINT = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "PAINT").Value);
		/** Standard Hyperlink / STD Moniker **/
		public static ClassIDPredefined STD_MONIKER = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "STD_MONIKER").Value);
		/** URL Moniker **/
		public static ClassIDPredefined URL_MONIKER = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "URL_MONIKER").Value);
		/** File Moniker **/
		public static ClassIDPredefined FILE_MONIKER = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "FILE_MONIKER").Value);
		/** Document summary information first property section **/
		public static ClassIDPredefined DOC_SUMMARY = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "DOC_SUMMARY").Value);
		/** Document summary information user defined properties section **/
		public static ClassIDPredefined USER_PROPERTIES = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "USER_PROPERTIES").Value);
		/** Summary information property section **/
		public static ClassIDPredefined SUMMARY_PROPERTIES = GetInstance(ClassIDPredefinedLookup.SingleOrDefault(c => c.Key == "SUMMARY_PROPERTIES").Value);

		private string externalForm;
		private ClassID classId;
		private string fileExtension;
		private string contentType;
		
		public ClassIDPredefined(string externalForm, string fileExtension, string contentType) 
		{
			this.externalForm = externalForm;
			this.fileExtension = fileExtension;
			this.contentType = contentType;
		}

		public static ClassIDPredefined GetInstance((string externalForm, string fileExtension, string contentType) value)
		{
			return new ClassIDPredefined(value.externalForm, value.fileExtension, value.contentType);
		}

		public ClassID getClassID()
		{
			lock(this) 
			{
				// TODO: init classId directly in the constructor when old statics have been removed from ClassID
				if (classId == null)
				{
					classId = new ClassID(externalForm);
				}
			}
			return classId;
		}
		
		public String getFileExtension()
		{
			return fileExtension;
		}
		
		public String getContentType()
		{
			return contentType;
		}
		
		public static ClassIDPredefined lookup(string externalForm)
		{
			var ClassIDEnum = ClassIDPredefinedLookup.SingleOrDefault(c => c.Value.externalForm == externalForm).Value;
			return new ClassIDPredefined(ClassIDEnum.externalForm, ClassIDEnum.fileExtension, ClassIDEnum.contentType);
		}
		
		public static ClassIDPredefined lookup(ClassID classID)
		{
			return (classID == null) ? null : lookup(classID.ToString());
		}

		//@SuppressWarnings("java:S1201")
		public bool equals(ClassID classID)
		{
			return getClassID().Equals(classID);
		}
	}
}
