using System;
using NPOI.HSSF.Record.AutoFilter;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;

namespace NPOI.HSSF.UserModel
{
    public class HSSFAutoFilter : IAutoFilter
    {
        // fix warning CS0649: Field 'autofilterinfo' is never assigned to, and will always have its default value null
        // AutoFilterInfoRecord autofilterinfo;
        // fix warning CS0169 "never used": AutoFilterRecord autofilter;
        FilterModeRecord filtermode;

        private HSSFSheet _sheet;

        public HSSFAutoFilter(HSSFSheet sheet)
        {
            _sheet = sheet;
        }

        public HSSFAutoFilter(string formula,HSSFWorkbook workbook)
        {
            //this.workbook = workbook;

            Ptg[] ptgs = HSSFFormulaParser.Parse(formula, workbook);
            if (!(ptgs[0] is Area3DPtg))
                throw new ArgumentException("incorrect formula");

            Area3DPtg ptg = (Area3DPtg)ptgs[0];
            HSSFSheet sheet = (HSSFSheet)workbook.GetSheetAt(ptg.ExternSheetIndex);
            //look for the prior record
            int loc = sheet.Sheet.FindFirstRecordLocBySid(DefaultColWidthRecord.sid) ;
            CreateFilterModeRecord(sheet, loc+1);
            CreateAutoFilterInfoRecord(sheet, loc + 2,ptg);
            //look for "_FilterDatabase" NameRecord of the sheet
            NameRecord name = workbook.Workbook.GetSpecificBuiltinRecord(NameRecord.BUILTIN_FILTER_DB, ptg.ExternSheetIndex+1);
            if (name == null)
                name = workbook.Workbook.CreateBuiltInName(NameRecord.BUILTIN_FILTER_DB, ptg.ExternSheetIndex + 1);
            name.IsHiddenName = true;

            name.NameDefinition = ptgs;
        }

        private void CreateFilterModeRecord(HSSFSheet sheet,int insertPos)
        {
            //look for the FilterModeRecord
            NPOI.HSSF.Record.Record record = sheet.Sheet.FindFirstRecordBySid(FilterModeRecord.sid);

            // this local variable hides the class one: FilterModeRecord filtermode;
            //if not found, add a new one
            if (record == null)
            {
                filtermode = new FilterModeRecord();
                sheet.Sheet.Records.Insert(insertPos, filtermode);
            }
        }

        private void CreateAutoFilterInfoRecord(HSSFSheet sheet, int insertPos, Area3DPtg ptg)
        {
            //look for the AutoFilterInfo Record
            NPOI.HSSF.Record.Record record = sheet.Sheet.FindFirstRecordBySid(AutoFilterInfoRecord.sid);
            AutoFilterInfoRecord info;
            if (record == null)
            {
                info = new AutoFilterInfoRecord();
                sheet.Sheet.Records.Insert(insertPos, info);
            }
            else
            {
                info = record as AutoFilterInfoRecord;
            }
            info.NumEntries = (short)(ptg.LastColumn - ptg.FirstColumn + 1);
        }

        private void RemoveFilterModeRecord(HSSFSheet sheet)
        {
            if(filtermode!=null)
                sheet.Sheet.Records.Remove(filtermode);
        }
        // fix warning autofilterinfo "is never assigned":
        //private void RemoveAutoFilterInfoRecord(HSSFSheet sheet)
        //{
        //    if (autofilterinfo != null)
        //        sheet.Sheet.Records.Remove(autofilterinfo);
        //}
    }
}
