using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.Model
{
    public class FIBCswNew
    {
        public short nFibNew=0;
        private short cQuickSavesNew=0;


        public FIBCswNew()
        {

        }

        public void Deserialize(HWPFStream stream)
        {
            nFibNew = stream.ReadShort();
            if (nFibNew == (short)0x0112)
            {
                ReadFibRgCswNewData2007(stream);
            }
            else
            {
                ReadFibRgCswNewData2000(stream);
            }            
        }

        private void ReadFibRgCswNewData2000(HWPFStream stream)
        {
            this.cQuickSavesNew = stream.ReadShort();
        }

        private void ReadFibRgCswNewData2007(HWPFStream stream)
        {
            ReadFibRgCswNewData2000(stream);
            stream.ReadShort();     //lidThemeOther
            stream.ReadShort();     //lidThemeFE
            stream.ReadShort();     //lidThemeCS
        }

        public short QuickSavesTimes
        {
            get { return cQuickSavesNew; }
            set { cQuickSavesNew = value; }
        }
    }
}
