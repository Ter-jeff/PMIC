//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

// 
// This source code was auto-generated by xsd, Version=4.0.30319.33440.
// 

using System.Xml.Serialization;

namespace nWireDefinition.FileGeneratorMapper.FileGeneratorMapperConfig
{
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace="http://www.teradyne.com/FWTool/FileGeneratorMapper.xsd")]
    [XmlRoot(Namespace="http://www.teradyne.com/FWTool/FileGeneratorMapper.xsd", IsNullable=false)]
    public partial class Mappings {
        
        private Mapping[] mappingField;
        
        /// <remarks/>
        [XmlElement("Mapping")]
        public Mapping[] Mapping {
            get {
                return this.mappingField;
            }
            set {
                this.mappingField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace="http://www.teradyne.com/FWTool/FileGeneratorMapper.xsd")]
    public partial class Mapping {
        
        private string generationTypeField;
        
        private string[] generatorField;
        
        /// <remarks/>
        public string GenerationType {
            get {
                return this.generationTypeField;
            }
            set {
                this.generationTypeField = value;
            }
        }
        
        /// <remarks/>
        [XmlElement("Generator")]
        public string[] Generator {
            get {
                return this.generatorField;
            }
            set {
                this.generatorField = value;
            }
        }
    }
}
