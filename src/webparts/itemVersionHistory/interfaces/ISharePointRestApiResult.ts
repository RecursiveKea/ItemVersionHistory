export interface ISharePointListField {
    Hidden: boolean;
    InternalName: string;
    Title: string;
    TypeAsString: string;
    ReadOnlyField: boolean;
    SystemField: boolean;
    SchemaXml: string;
}

export interface ISharePointFileVersion {
    ID: number;
    Size: number;
}