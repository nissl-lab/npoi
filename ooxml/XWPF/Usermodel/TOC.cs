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
    using System.Text;


    public class TOC
    {

        CT_SdtBlock block;
        private bool isBuilt = false;

        public TOC()
            : this(new CT_SdtBlock())
        {
        }

        public TOC(CT_SdtBlock block)
        {
            this.block = block;
            CT_SdtPr sdtPr = block.AddNewSdtPr();
            CT_DecimalNumber id = sdtPr.AddNewId();
            id.val=("4844945");
            sdtPr.AddNewDocPartObj().AddNewDocPartGallery().val=("Table of Contents");
            CT_SdtEndPr sdtEndPr = block.AddNewSdtEndPr();
            CT_RPr rPr = sdtEndPr.AddNewRPr();
            CT_Fonts fonts = rPr.AddNewRFonts();
            fonts.asciiTheme=(ST_Theme.minorHAnsi);
            fonts.eastAsiaTheme = (ST_Theme.minorHAnsi);
            fonts.hAnsiTheme = (ST_Theme.minorHAnsi);
            fonts.cstheme = (ST_Theme.minorBidi);
            CT_SdtContentBlock content = block.AddNewSdtContent();
            CT_P p = content.AddNewP();
            byte[] b = Encoding.Unicode.GetBytes("00EF7E24");
            p.rsidR = b;
            p.rsidRDefault = b;
            CT_PPr pPr = p.AddNewPPr();
            pPr.AddNewPStyle().val = ("TOCHeading");
            pPr.AddNewJc().val = ST_Jc.center;
            CT_R run = p.AddNewR();
            run.AddNewRPr().AddNewSz().val = 48;
            run.AddNewT().Value = ("Table of Contents");
            run.AddNewBr().type = ST_BrType.textWrapping; // line break

            // TOC Field
            p = content.AddNewP();
            pPr = p.AddNewPPr();
            pPr.AddNewPStyle().val = "TOC1";
            pPr.AddNewRPr().AddNewNoProof();

            run = p.AddNewR();
            run.AddNewFldChar().fldCharType = ST_FldCharType.begin;

            run = p.AddNewR();
            CT_Text text = run.AddNewInstrText();
            text.space = "preserve";
            text.Value = (" TOC \\h \\z ");

            p.AddNewR().AddNewFldChar().fldCharType = ST_FldCharType.separate;

        }


        public CT_SdtBlock GetBlock()
        {
            return this.block;
        }

        public void AddRow(int level, String title, int page, String bookmarkRef)
        {
            CT_SdtContentBlock contentBlock = this.block.sdtContent;
            CT_P p = contentBlock.AddNewP();
            byte[] b = Encoding.Unicode.GetBytes("00EF7E24");
            p.rsidR = b;
            p.rsidRDefault = b;
            CT_PPr pPr = p.AddNewPPr();
            pPr.AddNewPStyle().val=("TOC" + level);
            CT_Tabs tabs = pPr.AddNewTabs();
            CT_TabStop tab = tabs.AddNewTab();
            tab.val=(ST_TabJc.right);
            tab.leader=(ST_TabTlc.dot);
            tab.pos = "8290"; //(new BigInteger("8290"));
            pPr.AddNewRPr().AddNewNoProof();
            CT_R Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.AddNewT().Value=(title);
            Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.AddNewTab();
            Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.AddNewFldChar().fldCharType=(ST_FldCharType.begin);
            // pageref run
            Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            CT_Text text = Run.AddNewInstrText();
            text.space = "preserve";// (Space.PRESERVE);
            // bookmark reference
            text.Value=(" PAGEREF _Toc" + bookmarkRef + " \\h ");
            p.AddNewR().AddNewRPr().AddNewNoProof();
            Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.AddNewFldChar().fldCharType=(ST_FldCharType.separate);
            // page number run
            Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.AddNewT().Value = page.ToString();
            Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.AddNewFldChar().fldCharType=(ST_FldCharType.end);
        }

        public CT_SdtBlock Build()
        {
            // append end field char for TOC - only once
            if (!isBuilt)
            {
                CT_SdtContentBlock contentBlock = block.sdtContent;
                CT_P p = contentBlock.AddNewP();
                CT_R run = p.AddNewR();
                run.AddNewRPr().AddNewNoProof();
                run.AddNewFldChar().fldCharType = ST_FldCharType.end;

                isBuilt = true;
            }

            return block;
        }
    }

}