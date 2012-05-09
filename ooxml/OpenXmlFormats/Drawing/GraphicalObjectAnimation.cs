using System.ComponentModel;
using System.Xml.Serialization;
namespace NPOI.OpenXmlFormats.Dml {
    
    

    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_ChartBuildStep {
        
    
        category,
        
    
        ptInCategory,
        
    
        series,
        
    
        ptInSeries,
        
    
        allPts,
        
    
        gridLegend,
    }
    

    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_DgmBuildStep {
        
    
        sp,
        
    
        bg,
    }
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AnimationDgmElement {
        
        private string idField;
        
        private ST_DgmBuildStep bldStepField;
        
        public CT_AnimationDgmElement() {
            this.idField = "{00000000-0000-0000-0000-000000000000}";
            this.bldStepField = ST_DgmBuildStep.sp;
        }
        
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
//        [System.Xml.Serialization.XmlAttributeAttribute(DataType = "token")]
        [DefaultValue("{00000000-0000-0000-0000-000000000000}")]
        public string id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(ST_DgmBuildStep.sp)]
        public ST_DgmBuildStep bldStep {
            get {
                return this.bldStepField;
            }
            set {
                this.bldStepField = value;
            }
        }
    }
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AnimationChartElement {
        
        private int seriesIdxField;
        
        private int categoryIdxField;
        
        private ST_ChartBuildStep bldStepField;
        
        public CT_AnimationChartElement() {
            this.seriesIdxField = -1;
            this.categoryIdxField = -1;
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(-1)]
        public int seriesIdx {
            get {
                return this.seriesIdxField;
            }
            set {
                this.seriesIdxField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(-1)]
        public int categoryIdx {
            get {
                return this.categoryIdxField;
            }
            set {
                this.categoryIdxField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        public ST_ChartBuildStep bldStep {
            get {
                return this.bldStepField;
            }
            set {
                this.bldStepField = value;
            }
        }
    }
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AnimationElementChoice {
        
        private object itemField;
        
    
        [System.Xml.Serialization.XmlElement("chart", typeof(CT_AnimationChartElement))]
        [System.Xml.Serialization.XmlElement("dgm", typeof(CT_AnimationDgmElement))]
        public object Item {
            get {
                return this.itemField;
            }
            set {
                this.itemField = value;
            }
        }
    }
    

    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_AnimationBuildType {
        
    
        allAtOnce,
    }
    

    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_AnimationDgmOnlyBuildType {
        
    
        one,
        
    
        lvlOne,
        
    
        lvlAtOnce,
    }
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AnimationDgmBuildProperties {
        
        private string bldField;
        
        private bool revField;
        
        public CT_AnimationDgmBuildProperties() {
            this.bldField = "allAtOnce";
            this.revField = false;
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue("allAtOnce")]
        public string bld {
            get {
                return this.bldField;
            }
            set {
                this.bldField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(false)]
        public bool rev {
            get {
                return this.revField;
            }
            set {
                this.revField = value;
            }
        }
    }
    

    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_AnimationChartOnlyBuildType {
        
    
        series,
        
    
        category,
        
    
        seriesEl,
        
    
        categoryEl,
    }
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AnimationChartBuildProperties {
        
        private string bldField;
        
        private bool animBgField;
        
        public CT_AnimationChartBuildProperties() {
            this.bldField = "allAtOnce";
            this.animBgField = true;
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue("allAtOnce")]
        public string bld {
            get {
                return this.bldField;
            }
            set {
                this.bldField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(true)]
        public bool animBg {
            get {
                return this.animBgField;
            }
            set {
                this.animBgField = value;
            }
        }
    }
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AnimationGraphicalObjectBuildProperties {
        
        private object itemField;
        
    
        [System.Xml.Serialization.XmlElement("bldChart", typeof(CT_AnimationChartBuildProperties))]
        [System.Xml.Serialization.XmlElement("bldDgm", typeof(CT_AnimationDgmBuildProperties))]
        public object Item {
            get {
                return this.itemField;
            }
            set {
                this.itemField = value;
            }
        }
    }
}
