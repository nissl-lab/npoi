using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Num
    {

        private CT_DecimalNumber abstractNumIdField;

        private List<CT_NumLvl> lvlOverrideField;

        private string numIdField;

        public CT_Num()
        {
            //this.lvlOverrideField = new List<CT_NumLvl>();
            //this.abstractNumIdField = new CT_DecimalNumber();
        }
        public static CT_Num Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Num ctObj = new CT_Num();
            ctObj.numId = XmlHelper.ReadString(node.Attributes["w:numId"]);
            ctObj.lvlOverride = new List<CT_NumLvl>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "abstractNumId")
                    ctObj.abstractNumId = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvlOverride")
                    ctObj.lvlOverride.Add(CT_NumLvl.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:numId", this.numId);
            sw.Write(">");
            if (this.abstractNumId != null)
                this.abstractNumId.Write(sw, "abstractNumId");
            if (this.lvlOverride != null)
            {
                foreach (CT_NumLvl x in this.lvlOverride)
                {
                    x.Write(sw, "lvlOverride");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }



        [XmlElement(Order = 0)]
        public CT_DecimalNumber abstractNumId
        {
            get
            {
                return this.abstractNumIdField;
            }
            set
            {
                this.abstractNumIdField = value;
            }
        }

        [XmlElement("lvlOverride", Order = 1)]
        public List<CT_NumLvl> lvlOverride
        {
            get
            {
                return this.lvlOverrideField;
            }
            set
            {
                this.lvlOverrideField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string numId
        {
            get
            {
                return this.numIdField;
            }
            set
            {
                this.numIdField = value;
            }
        }

        public CT_DecimalNumber AddNewAbstractNumId()
        {
            if (this.abstractNumIdField == null)
                abstractNumIdField = new CT_DecimalNumber();
            return abstractNumIdField;
        }

        public int SizeOfLvlOverrideArray()
        {
            return lvlOverrideField == null ? 0 : lvlOverrideField.Count;
        }

        public CT_NumLvl GetLvlOverrideArray(int i)
        {
            return lvlOverrideField == null ? null : lvlOverrideField[i];
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_NumLvl
    {

        private CT_DecimalNumber startOverrideField;

        private CT_Lvl lvlField;

        private string ilvlField;

        public CT_NumLvl()
        {
            //this.lvlField = new CT_Lvl();
            //this.startOverrideField = new CT_DecimalNumber();
        }
        public static CT_NumLvl Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_NumLvl ctObj = new CT_NumLvl();
            ctObj.ilvl = XmlHelper.ReadString(node.Attributes["w:ilvl"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "startOverride")
                    ctObj.startOverride = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvl")
                    ctObj.lvl = CT_Lvl.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:ilvl", this.ilvl);
            sw.Write(">");
            if (this.startOverride != null)
                this.startOverride.Write(sw, "startOverride");
            if (this.lvl != null)
                this.lvl.Write(sw, "lvl");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }



        [XmlElement(Order = 0)]
        public CT_DecimalNumber startOverride
        {
            get
            {
                return this.startOverrideField;
            }
            set
            {
                this.startOverrideField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_Lvl lvl
        {
            get
            {
                return this.lvlField;
            }
            set
            {
                this.lvlField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string ilvl
        {
            get
            {
                return this.ilvlField;
            }
            set
            {
                this.ilvlField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot("numbering", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_Numbering
    {

        private List<CT_NumPicBullet> numPicBulletField;

        private List<CT_AbstractNum> abstractNumField;

        private List<CT_Num> numField;

        private CT_DecimalNumber numIdMacAtCleanupField;
        public static CT_Numbering Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Numbering ctObj = new CT_Numbering();
            ctObj.numPicBullet = new List<CT_NumPicBullet>();
            ctObj.abstractNum = new List<CT_AbstractNum>();
            ctObj.num = new List<CT_Num>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "numIdMacAtCleanup")
                    ctObj.numIdMacAtCleanup = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "numPicBullet")
                    ctObj.numPicBullet.Add(CT_NumPicBullet.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "abstractNum")
                    ctObj.abstractNum.Add(CT_AbstractNum.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "num")
                    ctObj.num.Add(CT_Num.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw)
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write("<w:numbering xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" ");
            sw.Write("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" ");
            sw.Write("xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" ");
            sw.Write("xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ");
            sw.Write("xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">");
            if (this.numIdMacAtCleanup != null)
                this.numIdMacAtCleanup.Write(sw, "numIdMacAtCleanup");
            if (this.numPicBullet != null)
            {
                foreach (CT_NumPicBullet x in this.numPicBullet)
                {
                    x.Write(sw, "numPicBullet");
                }
            }
            if (this.abstractNum != null)
            {
                foreach (CT_AbstractNum x in this.abstractNum)
                {
                    x.Write(sw, "abstractNum");
                }
            }
            if (this.num != null)
            {
                foreach (CT_Num x in this.num)
                {
                    x.Write(sw, "num");
                }
            }
            sw.Write("</w:numbering>");
        }

        public CT_Numbering()
        {
            //this.numIdMacAtCleanupField = new CT_DecimalNumber();
            //this.numField = new List<CT_Num>();
            //this.abstractNumField = new List<CT_AbstractNum>();
            //this.numPicBulletField = new List<CT_NumPicBullet>();
        }

        [XmlElement("numPicBullet", Order = 0)]
        public List<CT_NumPicBullet> numPicBullet
        {
            get
            {
                return this.numPicBulletField;
            }
            set
            {
                this.numPicBulletField = value;
            }
        }

        [XmlElement("abstractNum", Order = 1)]
        public List<CT_AbstractNum> abstractNum
        {
            get
            {
                return this.abstractNumField;
            }
            set
            {
                this.abstractNumField = value;
            }
        }

        [XmlElement("num", Order = 2)]
        public List<CT_Num> num
        {
            get
            {
                return this.numField;
            }
            set
            {
                this.numField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_DecimalNumber numIdMacAtCleanup
        {
            get
            {
                return this.numIdMacAtCleanupField;
            }
            set
            {
                this.numIdMacAtCleanupField = value;
            }
        }

        public IList<CT_Num> GetNumList()
        {
            return numField;
        }

        public IList<CT_AbstractNum> GetAbstractNumList()
        {
            return abstractNumField;
        }

        public CT_Num AddNewNum()
        {
            CT_Num num = new CT_Num();
            if (this.numField == null)
                this.numField = new List<CT_Num>();
            numField.Add(num);
            return num;
        }

        public void SetNumArray(int pos, CT_Num ct_Num)
        {
            if (this.numField == null)
                this.numField = new List<CT_Num>();

            if (pos < 0 || pos >= numField.Count)
                numField.Add(ct_Num);
            numField[pos] = ct_Num;
        }

        public CT_AbstractNum AddNewAbstractNum()
        {
            CT_AbstractNum num = new CT_AbstractNum();
            if (this.abstractNumField == null)
                this.abstractNumField = new List<CT_AbstractNum>();
            this.abstractNumField.Add(num);
            return num;
        }

        public void SetAbstractNumArray(int pos, CT_AbstractNum cT_AbstractNum)
        {
            if (this.abstractNumField == null)
                this.abstractNumField = new List<CT_AbstractNum>();
            if (pos < 0 || pos >= abstractNumField.Count)
                abstractNumField.Add(cT_AbstractNum);
            abstractNumField[pos] = cT_AbstractNum;
        }

        public void RemoveAbstractNum(int p)
        {
            if (this.abstractNumField == null)
                return;
            abstractNumField.RemoveAt(p);
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_NumPicBullet
    {

        private CT_Picture pictField;

        private string numPicBulletIdField;

        public CT_NumPicBullet()
        {
            //this.pictField = new CT_Picture();
        }

        [XmlElement(Order = 0)]
        public CT_Picture pict
        {
            get
            {
                return this.pictField;
            }
            set
            {
                this.pictField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string numPicBulletId
        {
            get
            {
                return this.numPicBulletIdField;
            }
            set
            {
                this.numPicBulletIdField = value;
            }
        }
        public static CT_NumPicBullet Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_NumPicBullet ctObj = new CT_NumPicBullet();
            ctObj.numPicBulletId = XmlHelper.ReadString(node.Attributes["w:numPicBulletId"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "pict")
                    ctObj.pict = CT_Picture.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:numPicBulletId", this.numPicBulletId);
            sw.Write(">");
            if (this.pict != null)
                this.pict.Write(sw, "pict");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }




    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_NumFmt
    {
        public static CT_NumFmt Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_NumFmt ctObj = new CT_NumFmt();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_NumberFormat)Enum.Parse(typeof(ST_NumberFormat), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write("/>");
        }


        private ST_NumberFormat valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_NumberFormat val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_NumberFormat
    {

    
        [XmlEnum("decimal")]
        @decimal,

    
        upperRoman,

    
        lowerRoman,

    
        upperLetter,

    
        lowerLetter,

    
        ordinal,

    
        cardinalText,

    
        ordinalText,

    
        hex,

    
        chicago,

    
        ideographDigital,

    
        japaneseCounting,

    
        aiueo,

    
        iroha,

    
        decimalFullWidth,

    
        decimalHalfWidth,

    
        japaneseLegal,

    
        japaneseDigitalTenThousand,

    
        decimalEnclosedCircle,

    
        decimalFullWidth2,

    
        aiueoFullWidth,

    
        irohaFullWidth,

    
        decimalZero,

    
        bullet,

    
        ganada,

    
        chosung,

    
        decimalEnclosedFullstop,

    
        decimalEnclosedParen,

    
        decimalEnclosedCircleChinese,

    
        ideographEnclosedCircle,

    
        ideographTraditional,

    
        ideographZodiac,

    
        ideographZodiacTraditional,

    
        taiwaneseCounting,

    
        ideographLegalTraditional,

    
        taiwaneseCountingThousand,

    
        taiwaneseDigital,

    
        chineseCounting,

    
        chineseLegalSimplified,

    
        chineseCountingThousand,

    
        koreanDigital,

    
        koreanCounting,

    
        koreanLegal,

    
        koreanDigital2,

    
        vietnameseCounting,

    
        russianLower,

    
        russianUpper,

    
        none,

    
        numberInDash,

    
        hebrew1,

    
        hebrew2,

    
        arabicAlpha,

    
        arabicAbjad,

    
        hindiVowels,

    
        hindiConsonants,

    
        hindiNumbers,

    
        hindiCounting,

    
        thaiLetters,

    
        thaiNumbers,

    
        thaiCounting,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_NumRestart
    {
        public static CT_NumRestart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_NumRestart ctObj = new CT_NumRestart();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_RestartNumber)Enum.Parse(typeof(ST_RestartNumber), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        private ST_RestartNumber valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_RestartNumber val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_RestartNumber
    {

    
        continuous,

    
        eachSect,

    
        eachPage,
    }
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_NumPr
    {

        private CT_DecimalNumber ilvlField;

        private CT_DecimalNumber numIdField;

        private CT_TrackChangeNumbering numberingChangeField;

        private CT_TrackChange insField;

        public CT_NumPr()
        {
            //this.insField = new CT_TrackChange();
            //this.numberingChangeField = new CT_TrackChangeNumbering();
            //this.numIdField = new CT_DecimalNumber();
            //this.ilvlField = new CT_DecimalNumber();
        }
        public static CT_NumPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_NumPr ctObj = new CT_NumPr();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "ilvl")
                    ctObj.ilvl = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "numId")
                    ctObj.numId = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "numberingChange")
                    ctObj.numberingChange = CT_TrackChangeNumbering.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ins")
                    ctObj.ins = CT_TrackChange.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.ilvl != null)
                this.ilvl.Write(sw, "ilvl");
            if (this.numId != null)
                this.numId.Write(sw, "numId");
            if (this.numberingChange != null)
                this.numberingChange.Write(sw, "numberingChange");
            if (this.ins != null)
                this.ins.Write(sw, "ins");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }



        [XmlElement(Order = 0)]
        public CT_DecimalNumber ilvl
        {
            get
            {
                return this.ilvlField;
            }
            set
            {
                this.ilvlField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_DecimalNumber numId
        {
            get
            {
                return this.numIdField;
            }
            set
            {
                this.numIdField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_TrackChangeNumbering numberingChange
        {
            get
            {
                return this.numberingChangeField;
            }
            set
            {
                this.numberingChangeField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_TrackChange ins
        {
            get
            {
                return this.insField;
            }
            set
            {
                this.insField = value;
            }
        }

        public CT_DecimalNumber AddNewNumId()
        {
            this.numId = new CT_DecimalNumber();
            return this.numId;
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Sym
    {

        private string fontField;

        private byte[] charField;
        public static CT_Sym Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Sym ctObj = new CT_Sym();
            ctObj.font = XmlHelper.ReadString(node.Attributes["w:font"]);
            ctObj.@char = XmlHelper.ReadBytes(node.Attributes["w:char"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:font", this.font);
            XmlHelper.WriteAttribute(sw, "w:char", this.@char);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string font
        {
            get
            {
                return this.fontField;
            }
            set
            {
                this.fontField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] @char
        {
            get
            {
                return this.charField;
            }
            set
            {
                this.charField = value;
            }
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_AbstractNum
    {

        private CT_LongHexNumber nsidField;

        private CT_MultiLevelType multiLevelTypeField;

        private CT_LongHexNumber tmplField;

        private CT_String nameField;

        private CT_String styleLinkField;

        private CT_String numStyleLinkField;

        private List<CT_Lvl> lvlField;

        private string abstractNumIdField;

        public CT_AbstractNum()
        {
            //this.lvlField = new List<CT_Lvl>();
            //this.numStyleLinkField = new CT_String();
            //this.styleLinkField = new CT_String();
            //this.nameField = new CT_String();
            //this.tmplField = new CT_LongHexNumber();
            this.multiLevelTypeField = new CT_MultiLevelType();
            this.nsidField = new CT_LongHexNumber();
            this.nsidField.val = new byte[4];
            Array.Copy(BitConverter.GetBytes(DateTime.Now.Ticks), 4, this.nsidField.val, 0, 4);
        }

        [XmlElement(Order = 0)]
        public CT_LongHexNumber nsid
        {
            get
            {
                return this.nsidField;
            }
            set
            {
                this.nsidField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_MultiLevelType multiLevelType
        {
            get
            {
                return this.multiLevelTypeField;
            }
            set
            {
                this.multiLevelTypeField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_LongHexNumber tmpl
        {
            get
            {
                return this.tmplField;
            }
            set
            {
                this.tmplField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_String name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_String styleLink
        {
            get
            {
                return this.styleLinkField;
            }
            set
            {
                this.styleLinkField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_String numStyleLink
        {
            get
            {
                return this.numStyleLinkField;
            }
            set
            {
                this.numStyleLinkField = value;
            }
        }

        [XmlElement("lvl", Order = 6)]
        public List<CT_Lvl> lvl
        {
            get
            {
                return this.lvlField;
            }
            set
            {
                 this.lvlField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string abstractNumId
        {
            get
            {
                return this.abstractNumIdField;
            }
            set
            {
                this.abstractNumIdField = value;
            }
        }

        public CT_AbstractNum Copy()
        {
            CT_AbstractNum anum = new CT_AbstractNum();
            anum.abstractNumIdField = this.abstractNumIdField;
            anum.lvlField = new List<CT_Lvl>(this.lvlField);
            anum.multiLevelTypeField = this.multiLevelTypeField;
            anum.nameField = this.nameField;
            anum.nsidField = this.nsidField;
            anum.numStyleLinkField = this.numStyleLinkField;
            anum.styleLinkField = this.styleLinkField;
            anum.tmplField = this.tmplField;
            return anum;
        }

        public bool ValueEquals(CT_AbstractNum cT_AbstractNum)
        {
            return this.abstractNumIdField == cT_AbstractNum.abstractNumIdField;
        }

        public void Set(CT_AbstractNum cT_AbstractNum)
        {
            this.abstractNumIdField = cT_AbstractNum.abstractNumIdField;
            this.lvlField = new List<CT_Lvl>(cT_AbstractNum.lvlField);
            this.multiLevelTypeField = cT_AbstractNum.multiLevelTypeField;
            this.nameField = cT_AbstractNum.nameField;
            this.nsidField = cT_AbstractNum.nsidField;
            this.numStyleLinkField = cT_AbstractNum.numStyleLinkField;
            this.styleLinkField = cT_AbstractNum.styleLinkField;
            this.tmplField = cT_AbstractNum.tmplField;
        }

        public static CT_AbstractNum Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_AbstractNum ctObj = new CT_AbstractNum();
            ctObj.abstractNumId = XmlHelper.ReadString(node.Attributes["w:abstractNumId"]);
            ctObj.lvl = new List<CT_Lvl>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "nsid")
                    ctObj.nsid = CT_LongHexNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "multiLevelType")
                    ctObj.multiLevelType = CT_MultiLevelType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tmpl")
                    ctObj.tmpl = CT_LongHexNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "name")
                    ctObj.name = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "styleLink")
                    ctObj.styleLink = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "numStyleLink")
                    ctObj.numStyleLink = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvl")
                    ctObj.lvl.Add(CT_Lvl.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:abstractNumId", this.abstractNumId);
            sw.Write(">");
            if (this.nsid != null)
                this.nsid.Write(sw, "nsid");
            if (this.multiLevelType != null)
                this.multiLevelType.Write(sw, "multiLevelType");
            if (this.tmpl != null)
                this.tmpl.Write(sw, "tmpl");
            if (this.name != null)
                this.name.Write(sw, "name");
            if (this.styleLink != null)
                this.styleLink.Write(sw, "styleLink");
            if (this.numStyleLink != null)
                this.numStyleLink.Write(sw, "numStyleLink");
            if (this.lvl != null)
            {
                foreach (CT_Lvl x in this.lvl)
                {
                    x.Write(sw, "lvl");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }




        public int SizeOfLvlArray()
        {
            return this.lvlField == null ? 0 : this.lvlField.Count;
        }

        public CT_Lvl GetLvlArray(int i)
        {
            return this.lvlField == null ? null : this.lvlField[i];
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_MultiLevelType
    {
        public static CT_MultiLevelType Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MultiLevelType ctObj = new CT_MultiLevelType();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_MultiLevelType)Enum.Parse(typeof(ST_MultiLevelType), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write("/>");
        }

        private ST_MultiLevelType valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_MultiLevelType val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_MultiLevelType
    {

    
        singleLevel,

    
        multilevel,

    
        hybridMultilevel,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Lvl
    {

        private CT_DecimalNumber startField;

        private CT_NumFmt numFmtField;

        private CT_DecimalNumber lvlRestartField;

        private CT_String pStyleField;

        private CT_OnOff isLglField;

        private CT_LevelSuffix suffField;

        private CT_LevelText lvlTextField;

        private CT_DecimalNumber lvlPicBulletIdField;

        private CT_LvlLegacy legacyField;

        private CT_Jc lvlJcField;

        private CT_PPr pPrField;

        private CT_RPr rPrField;

        private string ilvlField;

        private byte[] tplcField;

        private ST_OnOff tentativeField;

        private bool tentativeFieldSpecified;

        public CT_Lvl()
        {
            this.rPrField = new CT_RPr();
            this.pPrField = new CT_PPr();
            this.lvlJcField = new CT_Jc();
            //this.legacyField = new CT_LvlLegacy();
            //this.lvlPicBulletIdField = new CT_DecimalNumber();
            this.lvlTextField = new CT_LevelText();
            //this.suffField = new CT_LevelSuffix();
            //this.isLglField = new CT_OnOff();
            //this.pStyleField = new CT_String();
            //this.lvlRestartField = new CT_DecimalNumber();
            this.numFmtField = new CT_NumFmt();
            this.startField = new CT_DecimalNumber();
        }
        public static CT_Lvl Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Lvl ctObj = new CT_Lvl();
            ctObj.ilvl = XmlHelper.ReadString(node.Attributes["w:ilvl"]);
            ctObj.tplc = XmlHelper.ReadBytes(node.Attributes["w:tplc"]);
            if (node.Attributes["w:tentative"] != null)
                ctObj.tentative = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:tentative"].Value);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "start")
                    ctObj.start = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "numFmt")
                    ctObj.numFmt = CT_NumFmt.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvlRestart")
                    ctObj.lvlRestart = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pStyle")
                    ctObj.pStyle = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "isLgl")
                    ctObj.isLgl = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "suff")
                    ctObj.suff = CT_LevelSuffix.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvlText")
                    ctObj.lvlText = CT_LevelText.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvlPicBulletId")
                    ctObj.lvlPicBulletId = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "legacy")
                    ctObj.legacy = CT_LvlLegacy.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvlJc")
                    ctObj.lvlJc = CT_Jc.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pPr")
                    ctObj.pPr = CT_PPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rPr")
                    ctObj.rPr = CT_RPr.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:ilvl", this.ilvl);
            XmlHelper.WriteAttribute(sw, "w:tplc", this.tplc);
            if(this.tentative!= ST_OnOff.off)
                XmlHelper.WriteAttribute(sw, "w:tentative", this.tentative.ToString());
            sw.Write(">");
            if (this.start != null)
                this.start.Write(sw, "start");
            if (this.numFmt != null)
                this.numFmt.Write(sw, "numFmt");
            if (this.lvlRestart != null)
                this.lvlRestart.Write(sw, "lvlRestart");
            if (this.pStyle != null)
                this.pStyle.Write(sw, "pStyle");
            if (this.isLgl != null)
                this.isLgl.Write(sw, "isLgl");
            if (this.suff != null)
                this.suff.Write(sw, "suff");
            if (this.lvlText != null)
                this.lvlText.Write(sw, "lvlText");
            if (this.lvlPicBulletId != null)
                this.lvlPicBulletId.Write(sw, "lvlPicBulletId");
            if (this.legacy != null)
                this.legacy.Write(sw, "legacy");
            if (this.lvlJc != null)
                this.lvlJc.Write(sw, "lvlJc");
            if (this.pPr != null)
                this.pPr.Write(sw, "pPr");
            if (this.rPr != null)
                this.rPr.Write(sw, "rPr");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }



        [XmlElement(Order = 0)]
        public CT_DecimalNumber start
        {
            get
            {
                return this.startField;
            }
            set
            {
                this.startField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_NumFmt numFmt
        {
            get
            {
                return this.numFmtField;
            }
            set
            {
                this.numFmtField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_DecimalNumber lvlRestart
        {
            get
            {
                return this.lvlRestartField;
            }
            set
            {
                this.lvlRestartField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_String pStyle
        {
            get
            {
                return this.pStyleField;
            }
            set
            {
                this.pStyleField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_OnOff isLgl
        {
            get
            {
                return this.isLglField;
            }
            set
            {
                this.isLglField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_LevelSuffix suff
        {
            get
            {
                return this.suffField;
            }
            set
            {
                this.suffField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_LevelText lvlText
        {
            get
            {
                return this.lvlTextField;
            }
            set
            {
                this.lvlTextField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_DecimalNumber lvlPicBulletId
        {
            get
            {
                return this.lvlPicBulletIdField;
            }
            set
            {
                this.lvlPicBulletIdField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_LvlLegacy legacy
        {
            get
            {
                return this.legacyField;
            }
            set
            {
                this.legacyField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_Jc lvlJc
        {
            get
            {
                return this.lvlJcField;
            }
            set
            {
                this.lvlJcField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_PPr pPr
        {
            get
            {
                return this.pPrField;
            }
            set
            {
                this.pPrField = value;
            }
        }

        [XmlElement(Order = 11)]
        public CT_RPr rPr
        {
            get
            {
                return this.rPrField;
            }
            set
            {
                this.rPrField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string ilvl
        {
            get
            {
                return this.ilvlField;
            }
            set
            {
                this.ilvlField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] tplc
        {
            get
            {
                return this.tplcField;
            }
            set
            {
                this.tplcField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff tentative
        {
            get
            {
                return this.tentativeField;
            }
            set
            {
                this.tentativeField = value;
                this.tentativeFieldSpecified = true;
            }
        }

        [XmlIgnore]
        public bool tentativeSpecified
        {
            get
            {
                return this.tentativeFieldSpecified;
            }
            set
            {
                this.tentativeFieldSpecified = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_LevelSuffix
    {
        public static CT_LevelSuffix Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_LevelSuffix ctObj = new CT_LevelSuffix();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_LevelSuffix)Enum.Parse(typeof(ST_LevelSuffix), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write("/>");
        }

        private ST_LevelSuffix valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_LevelSuffix val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_LevelSuffix
    {
        tab,
        space,
        nothing,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_LevelText
    {

        private string valField;

        private ST_OnOff nullField= ST_OnOff.off;

        private bool nullFieldSpecified;

        public static CT_LevelText Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if(node==null)
                return null;
            CT_LevelText ctObj = new CT_LevelText();
            ctObj.val = XmlHelper.ReadString(node.Attributes["w:val"]);
            if (node.Attributes["w:null"]!=null)
                ctObj.@null = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:null"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val);
            if(this.@null!= ST_OnOff.off)
                XmlHelper.WriteAttribute(sw, "w:null", this.@null.ToString());
            sw.Write("/>");
        }


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff @null
        {
            get
            {
                return this.nullField;
            }
            set
            {
                this.nullField = value;
            }
        }

        [XmlIgnore]
        public bool nullSpecified
        {
            get
            {
                return this.nullFieldSpecified;
            }
            set
            {
                this.nullFieldSpecified = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_LvlLegacy
    {

        private ST_OnOff legacyField;

        private bool legacyFieldSpecified;

        private ulong legacySpaceField;

        private bool legacySpaceFieldSpecified;

        private string legacyIndentField;

        public static CT_LvlLegacy Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_LvlLegacy ctObj = new CT_LvlLegacy();
            if (node.Attributes["w:legacy"] != null)
                ctObj.legacy = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:legacy"].Value);
            ctObj.legacySpace = XmlHelper.ReadULong(node.Attributes["w:legacySpace"]);
            ctObj.legacyIndent = XmlHelper.ReadString(node.Attributes["w:legacyIndent"]);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:legacy", this.legacy.ToString());
            XmlHelper.WriteAttribute(sw, "w:legacySpace", this.legacySpace);
            XmlHelper.WriteAttribute(sw, "w:legacyIndent", this.legacyIndent);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff legacy
        {
            get
            {
                return this.legacyField;
            }
            set
            {
                this.legacyField = value;
            }
        }

        [XmlIgnore]
        public bool legacySpecified
        {
            get
            {
                return this.legacyFieldSpecified;
            }
            set
            {
                this.legacyFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong legacySpace
        {
            get
            {
                return this.legacySpaceField;
            }
            set
            {
                this.legacySpaceField = value;
            }
        }

        [XmlIgnore]
        public bool legacySpaceSpecified
        {
            get
            {
                return this.legacySpaceFieldSpecified;
            }
            set
            {
                this.legacySpaceFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string legacyIndent
        {
            get
            {
                return this.legacyIndentField;
            }
            set
            {
                this.legacyIndentField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_LsdException
    {

        private string nameField;

        private ST_OnOff lockedField;

        private bool lockedFieldSpecified;

        private string uiPriorityField;

        private ST_OnOff semiHiddenField;

        private bool semiHiddenFieldSpecified;

        private ST_OnOff unhideWhenUsedField;

        private bool unhideWhenUsedFieldSpecified;

        private ST_OnOff qFormatField;

        private bool qFormatFieldSpecified;
        public CT_LsdException()
        {
            semiHidden = ST_OnOff.off;
            unhideWhenUsed = ST_OnOff.off;
            locked = ST_OnOff.off;
        }

        public static CT_LsdException Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_LsdException ctObj = new CT_LsdException();
            ctObj.name = XmlHelper.ReadString(node.Attributes["w:name"]);
            if (node.Attributes["w:locked"] != null)
                ctObj.locked = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:locked"].Value);
            ctObj.uiPriority = XmlHelper.ReadString(node.Attributes["w:uiPriority"]);
            if (node.Attributes["w:semiHidden"] != null)
                ctObj.semiHidden = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:semiHidden"].Value);
            if (node.Attributes["w:unhideWhenUsed"] != null)
                ctObj.unhideWhenUsed = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:unhideWhenUsed"].Value);
            if (node.Attributes["w:qFormat"] != null)
                ctObj.qFormat = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:qFormat"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:name", this.name);
            if (locked != ST_OnOff.off)
                XmlHelper.WriteAttribute(sw, "w:locked", this.locked.ToString());
            if(this.semiHidden== ST_OnOff.on)
                XmlHelper.WriteAttribute(sw, "w:semiHidden", "1");
            XmlHelper.WriteAttribute(sw, "w:uiPriority", this.uiPriority);
            if (this.unhideWhenUsed == ST_OnOff.on)
                XmlHelper.WriteAttribute(sw, "w:unhideWhenUsed", "1");
            if (qFormat != ST_OnOff.off)
                XmlHelper.WriteAttribute(sw, "w:qFormat", this.qFormat.ToString());
            sw.Write("/>");
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff locked
        {
            get
            {
                return this.lockedField;
            }
            set
            {
                this.lockedField = value;
            }
        }

        [XmlIgnore]
        public bool lockedSpecified
        {
            get
            {
                return this.lockedFieldSpecified;
            }
            set
            {
                this.lockedFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string uiPriority
        {
            get
            {
                return this.uiPriorityField;
            }
            set
            {
                this.uiPriorityField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff semiHidden
        {
            get
            {
                return this.semiHiddenField;
            }
            set
            {
                this.semiHiddenField = value;
            }
        }

        [XmlIgnore]
        public bool semiHiddenSpecified
        {
            get
            {
                return this.semiHiddenFieldSpecified;
            }
            set
            {
                this.semiHiddenFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff unhideWhenUsed
        {
            get
            {
                return this.unhideWhenUsedField;
            }
            set
            {
                this.unhideWhenUsedField = value;
            }
        }

        [XmlIgnore]
        public bool unhideWhenUsedSpecified
        {
            get
            {
                return this.unhideWhenUsedFieldSpecified;
            }
            set
            {
                this.unhideWhenUsedFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff qFormat
        {
            get
            {
                return this.qFormatField;
            }
            set
            {
                this.qFormatField = value;
            }
        }

        [XmlIgnore]
        public bool qFormatSpecified
        {
            get
            {
                return this.qFormatFieldSpecified;
            }
            set
            {
                this.qFormatFieldSpecified = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TrackChangeNumbering : CT_TrackChange
    {
        public static new CT_TrackChangeNumbering Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TrackChangeNumbering ctObj = new CT_TrackChangeNumbering();
            ctObj.original = XmlHelper.ReadString(node.Attributes["original"]);
            ctObj.author = XmlHelper.ReadString(node.Attributes["author"]);
            ctObj.date = XmlHelper.ReadString(node.Attributes["date"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["id"]);
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "original", this.original);
            XmlHelper.WriteAttribute(sw, "author", this.author);
            XmlHelper.WriteAttribute(sw, "date", this.date);
            XmlHelper.WriteAttribute(sw, "id", this.id);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        private string originalField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string original
        {
            get
            {
                return this.originalField;
            }
            set
            {
                this.originalField = value;
            }
        }
    }

}
