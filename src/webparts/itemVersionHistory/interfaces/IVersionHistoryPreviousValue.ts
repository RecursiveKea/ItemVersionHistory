import { ISharePointListField } from './ISharePointRestApiResult';

export interface IVersionHistoryPreviousValue {
    OldValue: string;
    OldValueJSON: string;
    OldVersionID: number;
    OldVersionLabel: string;    
}

export interface IVersionHistoryPreviousValueChange extends IVersionHistoryPreviousValue {
  OldVersionValueNumberRange: string;
  NewValue: string;
  NewValueJSON: string;
  Field: ISharePointListField;  
}

export interface IVersionHistoryChange {
  Change: IVersionHistoryPreviousValueChange;
  VersionID: number;
}

export interface IItemVersionHistoryFieldValue {
  FieldIntenalName:string;
  NewValue: string;
  NewValueJSON: string;
  IsChangeFromPreviousVersion: boolean;
}

export interface IItemVersionHistory {
  VersionID: number;
  VersionLabel: string;
  VersionUrl: string;
  VersionFileSize: string;
  VersionModifiedBy: string;
  VersionModifiedByUrl: string;
  VersionModifiedDate: string;
  VersionMetadata: Map<string, IItemVersionHistoryFieldValue>;
  VersionMetadataChanges: Map<string, IVersionHistoryPreviousValueChange>;
}