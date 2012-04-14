/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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
namespace NPOI.XWPF.UserModel
{
    using System;
    using NPOI.OpenXmlFormats.Wordprocessing;


    public class TOC
    {

        CT_SdtBlock block;

        public TOC()
            : this(new CT_SdtBlock())
        {
        }

        public TOC(CT_SdtBlock block)
        {
            /*this.block = block;
            CT_SdtPr sdtPr = block.AddNewSdtPr();
            CT_DecimalNumber id = sdtPr.AddNewId();
            id.Val=(new Bigint("4844945"));
            sdtPr.AddNewDocPartObj().AddNewDocPartGallery().Val=("Table of contents");
            CTSdtEndPr sdtEndPr = block.AddNewSdtEndPr();
            CTRPr rPr = sdtEndPr.AddNewRPr();
            CTFonts fonts = rPr.AddNewRFonts();
            fonts.AsciiTheme=(STTheme.MINOR_H_ANSI);
            fonts.EastAsiaTheme=(STTheme.MINOR_H_ANSI);
            fonts.HAnsiTheme=(STTheme.MINOR_H_ANSI);
            fonts.Cstheme=(STTheme.MINOR_BIDI);
            rPr.AddNewB().Val=(STOnOff.OFF);
            rPr.AddNewBCs().Val=(STOnOff.OFF);
            rPr.AddNewColor().Val=("auto");
            rPr.AddNewSz().Val=(new Bigint("24"));
            rPr.AddNewSzCs().Val=(new Bigint("24"));
            CTSdtContentBlock content = block.AddNewSdtContent();
            CTP p = content.AddNewP();
            p.RsidR=("00EF7E24".Bytes);
            p.RsidRDefault=("00EF7E24".Bytes);
            p.AddNewPPr().AddNewPStyle().Val=("TOCHeading");
            p.AddNewR().AddNewT().StringValue=("Table of Contents");
             */
            throw new NotImplementedException();
        }


        public CT_SdtBlock GetBlock()
        {
            return this.block;
        }

        public void AddRow(int level, String title, int page, String bookmarkRef)
        {
            throw new NotImplementedException();
            /*CTSdtContentBlock contentBlock = this.block.SdtContent;
            CTP p = contentBlock.AddNewP();
            p.RsidR=("00EF7E24".Bytes);
            p.RsidRDefault=("00EF7E24".Bytes);
            CTPPr pPr = p.AddNewPPr();
            pPr.AddNewPStyle().Val=("TOC" + level);
            CTTabs tabs = pPr.AddNewTabs();
            CTTabStop tab = tabs.AddNewTab();
            tab.Val=(STTabJc.RIGHT);
            tab.Leader=(STTabTlc.DOT);
            tab.Pos=(new Bigint("8290"));
            pPr.AddNewRPr().AddNewNoProof();
            CTR run = p.AddNewR();
            Run.AddNewRPr().addNewNoProof();
            Run.AddNewT().StringValue=(title);
            run = p.AddNewR();
            Run.AddNewRPr().addNewNoProof();
            Run.AddNewTab();
            run = p.AddNewR();
            Run.AddNewRPr().addNewNoProof();
            Run.AddNewFldChar().FldCharType=(STFldCharType.BEGIN);
            // pageref run
            run = p.AddNewR();
            Run.AddNewRPr().addNewNoProof();
            CTText text = Run.AddNewInstrText();
            text.Space=(Space.PRESERVE);
            // bookmark reference
            text.StringValue=(" PAGEREF _Toc" + bookmarkRef + " \\h ");
            p.AddNewR().AddNewRPr().addNewNoProof();
            run = p.AddNewR();
            Run.AddNewRPr().addNewNoProof();
            Run.AddNewFldChar().FldCharType=(STFldCharType.SEPARATE);
            // page number run
            run = p.AddNewR();
            Run.AddNewRPr().addNewNoProof();
            Run.AddNewT().StringValue=(Int32.ValueOf(page).ToString());
            run = p.AddNewR();
            Run.AddNewRPr().addNewNoProof();
            Run.AddNewFldChar().FldCharType=(STFldCharType.END);*/
        }
    }

}