namespace NPOI.OpenXmlFormats.Dml {
    
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_ChartBuildStep {
        
        /// <remarks/>
        category,
        
        /// <remarks/>
        ptInCategory,
        
        /// <remarks/>
        series,
        
        /// <remarks/>
        ptInSeries,
        
        /// <remarks/>
        allPts,
        
        /// <remarks/>
        gridLegend,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_DgmBuildStep {
        
        /// <remarks/>
        sp,
        
        /// <remarks/>
        bg,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AnimationDgmElement {
        
        private string idField;
        
        private ST_DgmBuildStep bldStepField;
        
        public CT_AnimationDgmElement() {
            this.idField = "{00000000-0000-0000-0000-000000000000}";
            this.bldStepField = ST_DgmBuildStep.sp;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        [System.ComponentModel.DefaultValueAttribute("{00000000-0000-0000-0000-000000000000}")]
        public string id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_DgmBuildStep.sp)]
        public ST_DgmBuildStep bldStep {
            get {
                return this.bldStepField;
            }
            set {
                this.bldStepField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AnimationChartElement {
        
        private int seriesIdxField;
        
        private int categoryIdxField;
        
        private ST_ChartBuildStep bldStepField;
        
        public CT_AnimationChartElement() {
            this.seriesIdxField = -1;
            this.categoryIdxField = -1;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(-1)]
        public int seriesIdx {
            get {
                return this.seriesIdxField;
            }
            set {
                this.seriesIdxField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(-1)]
        public int categoryIdx {
            get {
                return this.categoryIdxField;
            }
            set {
                this.categoryIdxField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_ChartBuildStep bldStep {
            get {
                return this.bldStepField;
            }
            set {
                this.bldStepField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AnimationElementChoice {
        
        private object itemField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("chart", typeof(CT_AnimationChartElement))]
        [System.Xml.Serialization.XmlElementAttribute("dgm", typeof(CT_AnimationDgmElement))]
        public object Item {
            get {
                return this.itemField;
            }
            set {
                this.itemField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_AnimationBuildType {
        
        /// <remarks/>
        allAtOnce,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_AnimationDgmOnlyBuildType {
        
        /// <remarks/>
        one,
        
        /// <remarks/>
        lvlOne,
        
        /// <remarks/>
        lvlAtOnce,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AnimationDgmBuildProperties {
        
        private string bldField;
        
        private bool revField;
        
        public CT_AnimationDgmBuildProperties() {
            this.bldField = "allAtOnce";
            this.revField = false;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute("allAtOnce")]
        public string bld {
            get {
                return this.bldField;
            }
            set {
                this.bldField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool rev {
            get {
                return this.revField;
            }
            set {
                this.revField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_AnimationChartOnlyBuildType {
        
        /// <remarks/>
        series,
        
        /// <remarks/>
        category,
        
        /// <remarks/>
        seriesEl,
        
        /// <remarks/>
        categoryEl,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AnimationChartBuildProperties {
        
        private string bldField;
        
        private bool animBgField;
        
        public CT_AnimationChartBuildProperties() {
            this.bldField = "allAtOnce";
            this.animBgField = true;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute("allAtOnce")]
        public string bld {
            get {
                return this.bldField;
            }
            set {
                this.bldField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool animBg {
            get {
                return this.animBgField;
            }
            set {
                this.animBgField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AnimationGraphicalObjectBuildProperties {
        
        private object itemField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("bldChart", typeof(CT_AnimationChartBuildProperties))]
        [System.Xml.Serialization.XmlElementAttribute("bldDgm", typeof(CT_AnimationDgmBuildProperties))]
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
