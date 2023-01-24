using NPOI.HSLF.Model;
using NPOI.HSLF.Record;
using NPOI.SL.UserModel;
using System;

namespace NPOI.HSLF.UserModel
{
	public abstract class HSLFMasterSheet: HSLFSheet, MasterSheet<HSLFShape, HSLFTextParagraph>
	{
		public HSLFMasterSheet(SheetContainer container, int sheetNo)
			: base(container, sheetNo)
		{
			
		}
		public abstract TextPropCollection GetPropCollection(int txtype, int v, string pn, bool isChar);
		
	}
}