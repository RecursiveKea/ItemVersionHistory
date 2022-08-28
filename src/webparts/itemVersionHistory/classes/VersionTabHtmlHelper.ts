import { IVersionHistoryPreviousValue, IVersionHistoryPreviousValueChange, IItemVersionHistoryFieldValue, IItemVersionHistory, IVersionHistoryChange } from '../interfaces/IVersionHistoryPreviousValue';
import { ISharePointListField } from '../interfaces/ISharePointRestApiResult';
import { SharePointRestApiHelper } from '../CommonElements/classes/SharePointRestApiHelper';
import  styles from '../ItemVersionHistoryWebPart.module.scss'
import { HelperGenericFunctions } from '../CommonElements/classes/HelperGenericFunctions';
import { getIconClassName } from '@uifabric/styling';

export abstract class VersionTabHtmlHelper {

  public static WorkOutChanges(theSiteUrl: string, theFieldMapping: Map<string, ISharePointListField>, theVersions: Array<any>, theVersionFileSizes: Map<number, number>, theExtraInformation: boolean = false, isJSONComparison: boolean = false): Map<number, IItemVersionHistory> {

    const theSiteRelativeUrl: string = theSiteUrl.substring(theSiteUrl.indexOf(".sharepoint.com") + 15);
       
    const theFieldKeys: Array<string> = new Array<string>();
    theFieldMapping.forEach((value: ISharePointListField, key: string) => {
      theFieldKeys.push(key);
    });

    theFieldKeys.sort(HelperGenericFunctions.AscendingSort);

    const theVersionsKeys: Array<number> = new Array<number>();
    const theVersionsMap: Map<number, any> = new Map<number, any>();
    theVersions.forEach((currVersion: any) => {
      theVersionsMap.set(currVersion.VersionId, currVersion);
      theVersionsKeys.push(currVersion.VersionId);
    }); 

    theVersionsKeys.sort(HelperGenericFunctions.AscendingSort);

    const theResult: Map<number, IItemVersionHistory> = new Map<number, IItemVersionHistory>()
    const theOldValuesMap = new Map<string, IVersionHistoryPreviousValue>();
    let thePreviousVersionLabel: string = "";

    theVersionsKeys.forEach((currVersionId) => {
      const theCurrVersionData = theVersionsMap.get(currVersionId);      

      let theFileSizeText = "";
      if (theVersionFileSizes !== undefined) {
        const theFileSize: number = theVersionFileSizes.get(currVersionId);
        if (!HelperGenericFunctions.IsNullUndefinedOrEmpty(theFileSize)) {
          theFileSizeText = HelperGenericFunctions.GetFileSizeText(theFileSize);
        }
      }
      
      let theVersionUrl: string = SharePointRestApiHelper.GetListFieldValueFromRestResult("FileRef", "Text", theCurrVersionData);
      if (currVersionId !== theVersionsKeys[theVersionsKeys.length - 1]) {          
        theVersionUrl = theVersionUrl.substring(theVersionUrl.indexOf("/", theSiteRelativeUrl.length) + 1);
        theVersionUrl = (theSiteRelativeUrl + "/_vti_history/" + currVersionId.toString() + "/" + theVersionUrl);
      }

      const theVersionMetadata: Map<string, IItemVersionHistoryFieldValue> = new Map<string, IItemVersionHistoryFieldValue>();
      const theVersionMetadataChanges: Map<string, IVersionHistoryPreviousValueChange> = new Map<string, IVersionHistoryPreviousValueChange>();
      const theVersionLabel: string = SharePointRestApiHelper.GetListFieldValueFromRestResult("VersionLabel", "Text", theCurrVersionData);
      const theModifiedDate: string = SharePointRestApiHelper.GetListFieldValueFromRestResult("Modified", "DateTime", theCurrVersionData);
      const theModifiedBy: string = SharePointRestApiHelper.GetListFieldValueFromRestResult("Editor", "User", theCurrVersionData);
      let theModifiedByUserUrl: string = null;      

      if ((theCurrVersionData.Editor !== undefined) && (theCurrVersionData.Editor !== null)) {
        theModifiedByUserUrl = `${theSiteRelativeUrl}/_layouts/15/userdisp.aspx?ID=${theCurrVersionData.Editor.LookupId.toString()}`;
      }

      theFieldKeys.forEach((currField: string) => {
        const theField = theFieldMapping.get(currField);
        const theOldValueObj: IVersionHistoryPreviousValue = theOldValuesMap.get(currField);            
        let theOldValue = "(Blank)";
        let theOldValueJSON = "(Blank)";
        if (theOldValueObj !== undefined) {
          theOldValue = (theOldValueObj.OldValue);
          theOldValueJSON = theOldValueObj.OldValueJSON;
        }

        const theNewValue: string = SharePointRestApiHelper.GetListFieldValueFromRestResult(currField, theField.TypeAsString, theCurrVersionData, "(Blank)", theExtraInformation);
        const theNewValueJSON = (((theCurrVersionData[currField] === null)) ? "(Blank)" : JSON.stringify(theCurrVersionData[currField]));
        if ((theOldValueObj === undefined) || ((isJSONComparison) && (theOldValueJSON !== theNewValueJSON)) || (theOldValue !== theNewValue)) {

          let theOldVersionValueNumberRange: string = "-";
          if (theOldValueObj !== undefined) {
            if (theOldValueObj.OldVersionLabel === thePreviousVersionLabel) {
              theOldVersionValueNumberRange = (theOldValueObj.OldVersionLabel);
            } else {
              theOldVersionValueNumberRange = (theOldValueObj.OldVersionLabel + " - " + thePreviousVersionLabel);
            }
          }

          theVersionMetadataChanges.set(currField, {
            OldValue: ((theOldValueObj === undefined) ? "-" : theOldValueObj.OldValue),
            OldValueJSON: ((theOldValueObj === undefined) ? "-" : theOldValueObj.OldValueJSON),
            OldVersionID: ((theOldValueObj === undefined) ? null : theOldValueObj.OldVersionID),
            OldVersionLabel: ((theOldValueObj === undefined) ? null : theOldValueObj.OldVersionLabel),
            OldVersionValueNumberRange: theOldVersionValueNumberRange,
            NewValue: theNewValue,
            NewValueJSON: theNewValueJSON,
            Field: theField
          });

          theOldValuesMap.set(currField, {
            OldValue: theNewValue,
            OldValueJSON: theNewValueJSON,
            OldVersionID: currVersionId,
            OldVersionLabel: theVersionLabel
          });          
        }        

        theVersionMetadata.set(currField, {
          FieldIntenalName: currField,
          NewValue: theNewValue,
          NewValueJSON: theNewValueJSON,
          IsChangeFromPreviousVersion: ((theOldValueObj === undefined) || (theOldValue !== theNewValue))
        });
      });

      theResult.set(currVersionId, {
        VersionID: currVersionId,
        VersionLabel: theVersionLabel,
        VersionUrl: theVersionUrl,
        VersionFileSize: theFileSizeText,
        VersionModifiedBy: theModifiedBy,
        VersionModifiedByUrl: theModifiedByUserUrl,
        VersionModifiedDate: theModifiedDate,
        VersionMetadata: theVersionMetadata,
        VersionMetadataChanges: theVersionMetadataChanges
      });

      thePreviousVersionLabel = theVersionLabel;
    });

    return theResult;
  }

  private static GetModifiedByHtml(theVersion: IItemVersionHistory): string {
    let theResult = "";
    if (HelperGenericFunctions.IsNullUndefinedOrEmpty(theVersion.VersionModifiedByUrl)) {
      if (!HelperGenericFunctions.IsNullUndefinedOrEmpty(theVersion.VersionModifiedBy))
      theResult = (theVersion.VersionModifiedBy);
    } else {
      theResult = (`
        <a href="${theVersion.VersionModifiedByUrl}" target="_blank">
          ${theVersion.VersionModifiedBy}
        </a>
      `);
    }

    return theResult;
  }


  public static GetByVersionHistoryTabHtml(theVersions: Map<number, IItemVersionHistory>, showJSON: boolean = false): string {
    
    let theResult: string = "";
    
    const theVersionsKeys = HelperGenericFunctions.GetKeyArrayFromMap(theVersions);    
    theVersionsKeys.sort(HelperGenericFunctions.DescendingSort);
    theVersionsKeys.forEach((currVersionId) => {

      const currVersion: IItemVersionHistory = theVersions.get(currVersionId);

      let theCurrVersionFieldValuesHtml: string = "";
      theCurrVersionFieldValuesHtml += (`
        <div class='${styles.divByVersionWrapper}'>
          <div class='${styles.divVersionHeading}'>
            <div class='${styles.divVersionLabelModifiedBy}'>
              <a href="${currVersion.VersionUrl}" target="_blank">
                ${currVersion.VersionLabel}
              </a>
              :
              ${this.GetModifiedByHtml(currVersion)}
            </div>
            <div class='${styles.divVersionFileSize}'>
              ${currVersion.VersionFileSize}
            </div>
            <div class='${styles.divVersionModifiedDate}'>
              ${currVersion.VersionModifiedDate}
            </div>
          </div>
          <div class='${styles.divVersionFieldContent}'>                
      `);

      let theChangesHtml: string = "";
      currVersion.VersionMetadataChanges.forEach((value: IVersionHistoryPreviousValueChange, key: string) => {
        theChangesHtml += (`
            <tr>
              <td class="${styles.tdVersionHistoryValueListColumn}">
                <span title="${value.Field.InternalName}">
                  ${value.Field.Title}
                </span>
              </td>                  
              <td class="${styles.tdVersionHistoryValue}">
                <div title="${value.OldVersionValueNumberRange}">
                  ${((showJSON) ? value.OldValueJSON : value.OldValue)}
                </div>
              </td>
              <td class="${styles.tdVersionHistoryValue}">
                ${((showJSON) ? value.NewValueJSON : value.NewValue)}
              </td>
            </tr>
          `);
      });

      if (theChangesHtml !== "") {
        theCurrVersionFieldValuesHtml += (`
          <table>
            <tr>
              <th class="${styles.tdVersionHistoryValueListColumn}">Column</th>                  
              <th class="${styles.tdVersionHistoryValue}">Old Value</th>
              <th class="${styles.tdVersionHistoryValue}">New Value</th>
            </tr>
            ${theChangesHtml}
          </table>
        `);
      } else {
        theCurrVersionFieldValuesHtml += `<div class="${styles.divSpanNoChangesDetected}">No Metadata changes detected with filter criteria detected</div>`;
      }

      theCurrVersionFieldValuesHtml += "</div></div>";
      theResult += theCurrVersionFieldValuesHtml;      
    });

    return theResult;
  }

  public static GetByFieldHtml(theFieldMapping: Map<string, ISharePointListField>, theVersions: Map<number, IItemVersionHistory>, showJSON: boolean = false): string {

    const theFieldKeys = HelperGenericFunctions.GetKeyArrayFromMap(theFieldMapping);
    theFieldKeys.sort(HelperGenericFunctions.AscendingSort);

    const theVersionIDs = HelperGenericFunctions.GetKeyArrayFromMap(theVersions);
    theVersionIDs.sort(HelperGenericFunctions.AscendingSort);

    const theFieldChanges = new Map<string, Array<IVersionHistoryChange>>();
    theVersionIDs.forEach((currVersionID: number) => {
      
      const theCurrVersion = theVersions.get(currVersionID);

      theCurrVersion.VersionMetadataChanges.forEach((value: IVersionHistoryPreviousValueChange, key: string) => {
        const theFieldChangesArr = theFieldChanges.get(value.Field.InternalName);
        if (HelperGenericFunctions.IsNullUndefinedOrEmpty(theFieldChangesArr)) {
          theFieldChanges.set(value.Field.InternalName, new Array<IVersionHistoryChange>());
        }

        theFieldChanges.get(value.Field.InternalName).push({Change: value, VersionID: currVersionID});
      });
    });


    let theResult: string = "";
    theFieldKeys.forEach((currField: string) => {

      const theFieldChangesArr = theFieldChanges.get(currField);
      theFieldChangesArr.reverse();

      const theField = theFieldChangesArr[0].Change.Field;
      const isSystemField: boolean = (theField.SchemaXml.indexOf("http://schemas.microsoft.com/sharepoint/v3") !== -1);

      theResult += (`
        <div class="${styles.divByFieldWrapper}">
          <div class="${styles.divByFieldHeading}">            
            <div class="${styles.divByFieldHeadingText}">
              ${theField.InternalName} (Display Name: ${theField.Title}, Type: ${theField.TypeAsString})
            </div>
            <div class="${styles.divByFieldHeadingIcons}">
              ${((theField.Hidden) ?  "<i title='Hidden Field' class='" + getIconClassName("hide3") + "'></i>" : "" )}
              ${((theField.ReadOnlyField) ? "<i title='Read Only Field' class='" + getIconClassName("ReadingMode") + "'></i>" : "" )}
              ${((isSystemField) ? "<i title='System Field' class='" + getIconClassName("System") + "'></i>" : "" )}
            </div>
          </div>
          <div class="${styles.divByFieldContent}">
            <table>
              <tr>                
                <th class="${styles.tdByFieldSetByUser}">Set by</th>
                <th class="${styles.tdByFieldSetDate}">Set Date</th>
                <th class="${styles.tdByFieldVersions}">Versions</th>
                <th class="${styles.tdByFieldValue}">Value</th>
              </tr>
      `);
      
      const theLatestChangeVersion = theVersions.get(theFieldChangesArr[0].VersionID);
      let theVersionString: string = "";
      if ((theFieldChangesArr.length === 1) || (theFieldChangesArr[0].VersionID === theVersionIDs[theVersionIDs.length - 1])) {
        theVersionString = "Current";
      } else {
        theVersionString = (theLatestChangeVersion.VersionLabel + " - Current");
      }

      theResult += (`
        <tr>
          <td class="${styles.tdByFieldSetByUser}">${this.GetModifiedByHtml(theLatestChangeVersion)}</td>
          <td class="${styles.tdByFieldSetDate}">${theLatestChangeVersion.VersionModifiedDate}</td>
          <td class="${styles.tdByFieldVersions}">${theVersionString}</td>
          <td class="${styles.tdByFieldValue}">
            ${((showJSON) ? theFieldChangesArr[0].Change.NewValueJSON : theFieldChangesArr[0].Change.NewValue)}
          </td>
        </tr>
      `);

      if (theFieldChangesArr.length > 1) {
        theFieldChangesArr.forEach((currFieldValueChange: IVersionHistoryChange) => {          
          const theVerionWhereSet = theVersions.get(currFieldValueChange.Change.OldVersionID);

          if (!HelperGenericFunctions.IsNullUndefinedOrEmpty(theVerionWhereSet)) {
            theResult += (`
              <tr>
                <td>${this.GetModifiedByHtml(theVerionWhereSet)}</td>
                <td>${theVerionWhereSet.VersionModifiedDate}</td>
                <td>${currFieldValueChange.Change.OldVersionValueNumberRange}</td>
                <td>
                  ${((showJSON) ? currFieldValueChange.Change.OldValueJSON : currFieldValueChange.Change.OldValue)}
                </td>
              </tr>
            `);
          }          
        });
      }

      theResult += "</table></div></div>";
    });

    return theResult;
  }

  
  public static GetCompareVersionsHtml(theFieldMapping: Map<string, ISharePointListField>, theLeftVersion: IItemVersionHistory, theRightVersion: IItemVersionHistory, showJSON: boolean = false): string {
    
    const theFieldKeys: Array<string> = HelperGenericFunctions.GetKeyArrayFromMap(theFieldMapping);
    theFieldKeys.sort(HelperGenericFunctions.AscendingSort);

    let theResult: string = (`
      <table>
        <tr>
          <th class="${styles.tdCompareVersionField}">Field</th>
          <th class="${styles.tdCompareVersionComparison}">Comparison</th>
          <th class="${styles.tdCompareVersionFieldValue}">${theLeftVersion.VersionLabel}</th>
          <th class="${styles.tdCompareVersionFieldValue}">${theRightVersion.VersionLabel}</th>
        </tr>
    `);

    theFieldKeys.forEach((currField: string) => {
      const theField = theFieldMapping.get(currField);
      const theLeftFieldValue: string = ((showJSON) ? theLeftVersion.VersionMetadata.get(currField).NewValueJSON : theLeftVersion.VersionMetadata.get(currField).NewValue);
      const theRightFieldValue: string = ((showJSON) ? theRightVersion.VersionMetadata.get(currField).NewValueJSON : theRightVersion.VersionMetadata.get(currField).NewValue);
      
      theResult += (`
        <tr>
          <td class="${styles.tdCompareVersionField}">
            <span title="${theField.InternalName}">
              ${theField.Title}
            </span>
          </td>
          <td class="${styles.tdCompareVersionComparison}">${(theLeftFieldValue === theRightFieldValue) ? "Same": "Different"}</td>
          <td class="${styles.tdCompareVersionFieldValue}">${theLeftFieldValue}</td>
          <td class="${styles.tdCompareVersionFieldValue}">${theRightFieldValue}</td>
        </tr>
      `);
    });

    theResult += "</table>";
    return theResult;
  }

}