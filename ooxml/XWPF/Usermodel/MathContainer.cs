using NPOI.OpenXmlFormats.Shared;
using NPOI.XWPF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.XWPF.Usermodel
{
    public abstract class MathContainer : IRunBody
    {
        protected IRunBody parent;
        protected XWPFDocument document;
        protected IOMathContainer container;

        protected List<XWPFSharedRun> runs;
        protected List<XWPFNary> naries;
        protected List<XWPFAcc> accs;
        protected List<XWPFSSub> sSubs;
        protected List<XWPFSSup> sSups;
        protected List<XWPFF> fs;
        protected List<XWPFRad> rads;

        public MathContainer(IOMathContainer c, IRunBody p)
        {
            container = c;
            parent = p;
            document = p.Document;

            FillLists(c.Items);
        }

        public XWPFDocument Document
        {
            get { return document; }
        }

        public POIXMLDocumentPart Part { get { return parent.Part; } }

        private void FillLists(ArrayList items)
        {
            runs = new List<XWPFSharedRun>();
            accs = new List<XWPFAcc>();
            naries = new List<XWPFNary>();
            sSubs = new List<XWPFSSub>();
            sSups = new List<XWPFSSup>();
            fs = new List<XWPFF>();
            rads = new List<XWPFRad>();

            BuildListsInOrderFromXml(items);
        }

        private void BuildListsInOrderFromXml(ArrayList items)
        {
            foreach (object o in items)
            {
                if (o is CT_R r)
                {
                    runs.Add(new XWPFSharedRun(r, this));
                }
                if (o is CT_Acc acc)
                {
                    accs.Add(new XWPFAcc(acc, this));
                }

                if (o is CT_Nary nary)
                {
                    naries.Add(new XWPFNary(nary, this));
                }

                if (o is CT_SSub sub)
                {
                    sSubs.Add(new XWPFSSub(sub, this));
                }

                if (o is CT_F f)
                {
                    fs.Add(new XWPFF(f, this));
                }

                if (o is CT_Rad rad)
                {
                    rads.Add(new XWPFRad(rad, this));
                }
            }
        }

        public XWPFSharedRun CreateRun()
        {
            XWPFSharedRun run = new XWPFSharedRun(container.AddNewR(), this);
            runs.Add(run);
            return run;
        }

        /// <summary>
        /// Create Accent
        /// </summary>
        /// <returns></returns>
        public XWPFAcc CreateAcc()
        {
            XWPFAcc acc = new XWPFAcc(container.AddNewAcc(), this);
            accs.Add(acc);
            return acc;
        }

        /// <summary>
        /// Create n-ary Operator Object
        /// </summary>
        /// <returns></returns>
        public XWPFNary CreateNary()
        {
            XWPFNary nary = new XWPFNary(container.AddNewNary(), this);
            naries.Add(nary);
            return nary;
        }

        /// <summary>
        /// Subscript Object
        /// </summary>
        /// <returns></returns>
        public XWPFSSub CreateSSub()
        {
            XWPFSSub ssub = new XWPFSSub(container.AddNewSSub(), this);
            sSubs.Add(ssub);
            return ssub;
        }

        /// <summary>
        /// Superscript Object
        /// </summary>
        /// <returns></returns>
        public XWPFSSup CreateSSup()
        {
            XWPFSSup ssup = new XWPFSSup(container.AddNewSSup(), this);
            sSups.Add(ssup);
            return ssup;
        }

        /// <summary>
        /// Fraction Object
        /// </summary>
        /// <returns></returns>
        public XWPFF CreateF()
        {
            XWPFF f = new XWPFF(container.AddNewF(), this);
            fs.Add(f);
            return f;
        }

        /// <summary>
        /// Radical Object
        /// </summary>
        /// <returns></returns>
        public XWPFRad CreateRad()
        {
            XWPFRad rad = new XWPFRad(container.AddNewRad(), this);
            rads.Add(rad);
            return rad;
        }

        public IList<XWPFSharedRun> Runs
        {
            get
            {
                return runs.AsReadOnly();
            }
        }

        public IList<XWPFAcc> Accs
        {
            get
            {
                return accs.AsReadOnly();
            }
        }

        public IList<XWPFNary> Naries
        {
            get
            {
                return naries.AsReadOnly();
            }
        }

        public IList<XWPFSSub> SSubs
        {
            get
            {
                return sSubs.AsReadOnly();
            }
        }

        public IList<XWPFF> Fs
        {
            get
            {
                return fs.AsReadOnly();
            }
        }

        public IList<XWPFRad> Rads
        {
            get
            {
                return rads.AsReadOnly();
            }
        }


    }
}
