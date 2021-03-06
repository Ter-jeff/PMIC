//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using System.Xml.Serialization;

// 
// This source code was auto-generated by xsd, Version=4.0.30319.33440.
// 


/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true, Namespace="http://Teradyne.Oasis.IGData.Utilities/IGXLSheets.xsd")]
[System.Xml.Serialization.XmlRootAttribute(Namespace="http://Teradyne.Oasis.IGData.Utilities/IGXLSheets.xsd", IsNullable=false)]
public partial class IGXLVersion {
    
    private SheetInfo[] sheetsField;
    
    private string igxlVersionField;
    
    private IGXLVersionIgxlType igxlTypeField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlArrayItemAttribute("Sheet", IsNullable=false)]
    public SheetInfo[] Sheets {
        get {
            return this.sheetsField;
        }
        set {
            this.sheetsField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string igxlVersion {
        get {
            return this.igxlVersionField;
        }
        set {
            this.igxlVersionField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public IGXLVersionIgxlType igxlType {
        get {
            return this.igxlTypeField;
        }
        set {
            this.igxlTypeField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Teradyne.Oasis.IGData.Utilities/IGXLSheets.xsd")]
public partial class SheetInfo {
    
    private Field[] fieldField;
    
    private Columns columnsField;
    
    private string sheetNameField;
    
    private string sheetVersionField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute("Field")]
    public Field[] Field {
        get {
            return this.fieldField;
        }
        set {
            this.fieldField = value;
        }
    }
    
    /// <remarks/>
    public Columns Columns {
        get {
            return this.columnsField;
        }
        set {
            this.columnsField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string sheetName {
        get {
            return this.sheetNameField;
        }
        set {
            this.sheetNameField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string sheetVersion {
        get {
            return this.sheetVersionField;
        }
        set {
            this.sheetVersionField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Teradyne.Oasis.IGData.Utilities/IGXLSheets.xsd")]
public partial class Field {
    
    private string fieldNameField;
    
    private int rowIndexField;
    
    private int columnIndexField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string fieldName {
        get {
            return this.fieldNameField;
        }
        set {
            this.fieldNameField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public int rowIndex {
        get {
            return this.rowIndexField;
        }
        set {
            this.rowIndexField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public int columnIndex {
        get {
            return this.columnIndexField;
        }
        set {
            this.columnIndexField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Teradyne.Oasis.IGData.Utilities/IGXLSheets.xsd")]
public partial class Column {
    
    private Column[] column1Field;
    
    private string variantNameField;
    
    private string relativeColumnField;
    
    private string columnNameField;
    
    private bool isGroupField;
    
    private int indexFromField;
    
    private int indexToField;
    
    private int rowIndexField;
    
    public Column() {
        this.isGroupField = false;
        this.indexToField = 1;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute("Column")]
    public Column[] Column1 {
        get {
            return this.column1Field;
        }
        set {
            this.column1Field = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string variantName {
        get {
            return this.variantNameField;
        }
        set {
            this.variantNameField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string relativeColumn {
        get {
            return this.relativeColumnField;
        }
        set {
            this.relativeColumnField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string columnName {
        get {
            return this.columnNameField;
        }
        set {
            this.columnNameField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    [System.ComponentModel.DefaultValueAttribute(false)]
    public bool isGroup {
        get {
            return this.isGroupField;
        }
        set {
            this.isGroupField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public int indexFrom {
        get {
            return this.indexFromField;
        }
        set {
            this.indexFromField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    [System.ComponentModel.DefaultValueAttribute(1)]
    public int indexTo {
        get {
            return this.indexToField;
        }
        set {
            this.indexToField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public int rowIndex {
        get {
            return this.rowIndexField;
        }
        set {
            this.rowIndexField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Teradyne.Oasis.IGData.Utilities/IGXLSheets.xsd")]
public partial class Columns {
    
    private Column[] columnField;
    
    private Column[] variantField;
    
    private Column[] relativeColumnField;
    
    private int rowCountField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute("Column")]
    public Column[] Column {
        get {
            return this.columnField;
        }
        set {
            this.columnField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute("Variant")]
    public Column[] Variant {
        get {
            return this.variantField;
        }
        set {
            this.variantField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute("RelativeColumn")]
    public Column[] RelativeColumn {
        get {
            return this.relativeColumnField;
        }
        set {
            this.relativeColumnField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public int RowCount {
        get {
            return this.rowCountField;
        }
        set {
            this.rowCountField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true, Namespace="http://Teradyne.Oasis.IGData.Utilities/IGXLSheets.xsd")]
public enum IGXLVersionIgxlType {
    
    /// <remarks/>
    Flex,
    
    /// <remarks/>
    UltraFlex,
    
    /// <remarks/>
    J750,
    
    /// <remarks/>
    Generic,
}
