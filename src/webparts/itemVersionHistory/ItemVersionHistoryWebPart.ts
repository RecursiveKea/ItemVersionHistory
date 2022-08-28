import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ItemVersionHistoryWebPart.module.scss';
import CommonControlStyles from './CommonElements/stylesheets/CommonControlsFactoryStyle.module.scss';
import * as strings from 'ItemVersionHistoryWebPartStrings';

import { CommonControlsFactory } from './CommonElements/classes/CommonControlsFactory';
import { TabControl, TabController } from './CommonElements/classes/TabControl';
import { SharePointRestApiHelper } from './CommonElements/classes/SharePointRestApiHelper';
import { IRestApiResponse } from './CommonElements/interfaces/IRestApiResponse';
import { HelperGenericFunctions } from './CommonElements/classes/HelperGenericFunctions';
import { ISharePointListField, ISharePointFileVersion } from './interfaces/ISharePointRestApiResult';
import { VersionTabHtmlHelper } from './classes/VersionTabHtmlHelper';
import { IItemVersionHistory } from './interfaces/IVersionHistoryPreviousValue';


export interface IItemVersionHistoryWebPartProps {
  description: string;
}

export default class ItemVersionHistoryWebPart extends BaseClientSideWebPart<IItemVersionHistoryWebPartProps> {

  private readonly INPUT_CONTROL_FORM_ID: string = "divInputControls";
  private readonly RESULTS_TAB_CONTAINER_ID: string = "divResultsTabContainer";

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private renderInputForm(): void {

    const byUrlDiv: string = (`
      <div class="${ CommonControlStyles.divInputFlex }">
        
        ${CommonControlsFactory.getInputTextControlHtml("txtSiteUrl", "Site Url: ", 400)}
        ${CommonControlsFactory.getInputTextControlHtml("txtLibraryRootFolderPath", "List / Library Root Folder Path: ", 200)}
        ${CommonControlsFactory.getInputTextControlHtml("txtItemID", "Item ID: ", 50)}
        ${
          CommonControlsFactory.wrapLabelAndInputsInTemplateHtml("Display: ", `
            ${ 
              CommonControlsFactory.getInputControlWithLabelWithoutWrappingDivs("rbDisplayTypeText", "Text", "radio", true, "name='display_type' checked='checked' value='Text' ")
              + CommonControlsFactory.getInputControlWithLabelWithoutWrappingDivs("rbDisplayTypeJSON", "JSON", "radio", true, "name='display_type' value='JSON' ")
            }
          `)          
        }
        ${
          CommonControlsFactory.wrapLabelAndInputsInTemplateHtml("Comparison Type: ", `
            ${ 
              CommonControlsFactory.getInputControlWithLabelWithoutWrappingDivs("rbComparisonTypeText", "Text Comparison", "radio", true, "name='comparison_type' value='Text' ") 
              + CommonControlsFactory.getInputControlWithLabelWithoutWrappingDivs("rbComparisonTypeJSON", "JSON Comparison", "radio", true, "name='comparison_type'  checked='checked' value='JSON' ")
            }
          `)          
        }
        ${
          CommonControlsFactory.wrapLabelAndInputsInTemplateHtml("Options: ", `
            <div class="${CommonControlStyles.lblHover}">
            ${
              CommonControlsFactory.getInputControlWithLabelWithoutWrappingDivs("chkHiddenFields", "Hidden Fields", "checkbox", true)
              + "<br/>" + CommonControlsFactory.getInputControlWithLabelWithoutWrappingDivs("chkReadOnlyFields", "Read Only Fields", "checkbox", true)
              + "<br/>" + CommonControlsFactory.getInputControlWithLabelWithoutWrappingDivs("chkSystemFields", "System Fields", "checkbox", true)
              + "<br/>" + CommonControlsFactory.getInputControlWithLabelWithoutWrappingDivs("chkExtraInformation", "Extra Field Value Properties", "checkbox", true)
              + "<br/>" + CommonControlsFactory.getInputControlWithLabelWithoutWrappingDivs("chkExcludeVisibleFields", "Exclude Visible Fields", "checkbox", true)
            }
            </div>
          `)
        }
        ${CommonControlsFactory.getInputButtonControlHtml("btnVersionHistory", "Get Version History", styles.CustomButton)}
      </div>
    `);    

    document.getElementById(`${this.INPUT_CONTROL_FORM_ID}`).innerHTML = byUrlDiv;    
        
    const theHttpClient = this.context.spHttpClient;
    const theSearchButton: HTMLElement = document.getElementById("btnVersionHistory");
    const tabContainerID = this.RESULTS_TAB_CONTAINER_ID;

    theSearchButton.addEventListener("click", function(e) {

      let theSiteUrl: string = HelperGenericFunctions.GetTextFromHtmlElement("txtSiteUrl").trim();
      const theRootFolderPath: string = HelperGenericFunctions.GetTextFromHtmlElement("txtLibraryRootFolderPath").trim();
      const theItemID: string = HelperGenericFunctions.GetTextFromHtmlElement("txtItemID").trim();

      if (theSiteUrl.indexOf(".sharepoint.com") !== -1) {
        if (theRootFolderPath.length > 0) {
          if ((theItemID.length > 0) && (!isNaN(+theItemID))) {

            theSiteUrl = theSiteUrl.replace("%3A", ":");
            const theListRelativeUrl = ((theSiteUrl.substring(theSiteUrl.indexOf(".sharepoint.com") + 15)) + "/" + theRootFolderPath);
            const theListFieldsRestApiUrl: string = (theSiteUrl + "/_api/web/GetList('" + theListRelativeUrl + "')/fields?$select=Title,InternalName,ReadOnlyField,Hidden,TypeAsString,SchemaXml");
            const theVersionHistoryRestApiUrl: string = (theSiteUrl + "/_api/web/GetList('" + theListRelativeUrl + "')/items(" + theItemID + ")/versions");

            document.getElementById(tabContainerID).innerHTML = CommonControlsFactory.getLoadingHtml("Getting List/Library Fields...");
            
            SharePointRestApiHelper.RestApiCall(theHttpClient, theListFieldsRestApiUrl)
              .then((theListFields: IRestApiResponse) => {

                const theHiddenCheckBox: HTMLInputElement = document.getElementById("chkHiddenFields") as HTMLInputElement;
                const theReadOnlyCheckBox: HTMLInputElement = document.getElementById("chkReadOnlyFields") as HTMLInputElement;
                const theSystemFieldCheckBox: HTMLInputElement = document.getElementById("chkSystemFields") as HTMLInputElement;
                const theExcludeVisibleFieldsCheckBox: HTMLInputElement = document.getElementById("chkExcludeVisibleFields") as HTMLInputElement;                

                const theFieldsFormatted = new Map<string, ISharePointListField>();
                let theTitleField: ISharePointListField = null;
                let theFileLeafRefField: ISharePointListField = null;

                theListFields.value.forEach((currField: ISharePointListField) => {

                  const isSystemField = (currField.SchemaXml.indexOf("http://schemas.microsoft.com/sharepoint/v3") !== -1);

                  if (
                      ((!currField.ReadOnlyField) && (!currField.Hidden) && (!isSystemField) && (!theExcludeVisibleFieldsCheckBox.checked))
                      || ((currField.Hidden) && (theHiddenCheckBox.checked) && (!isSystemField))
                      || ((isSystemField) && (theSystemFieldCheckBox.checked) && (currField.TypeAsString !== "Computed"))
                      || ((currField.ReadOnlyField) && (theReadOnlyCheckBox.checked) && (!isSystemField) && (!currField.Hidden))
                    ) {
                      theFieldsFormatted.set(currField.InternalName, currField);
                  }

                  if (currField.InternalName === "Title") {
                    theTitleField = currField;
                  } else if (currField.InternalName === "FileLeafRef") {
                    theFileLeafRefField = currField;
                  }
                });

                if (theFieldsFormatted.size === 0) {
                  document.getElementById(tabContainerID).innerHTML = "No fields found with the options selected";
                } else {

                  if (theFieldsFormatted.get("Title") === undefined) {
                    theFieldsFormatted.set("Title", theTitleField);
                  }

                  if (theFieldsFormatted.get("FileLeafRef") === undefined) {
                    theFieldsFormatted.set("FileLeafRef", theFileLeafRefField);
                  }

                  document.getElementById(tabContainerID).innerHTML = CommonControlsFactory.getLoadingHtml("Getting Item Versions...");
                  
                  SharePointRestApiHelper.RestApiCall(theHttpClient, theVersionHistoryRestApiUrl)
                    .then((theVersionHistoryResponse: IRestApiResponse) => {  
                    
                      const theVersionHistory = (theVersionHistoryResponse.value);

                      const OutputResults = (theFileVersions?: Map<number, number>) => {
                        const theExtraInformation: HTMLInputElement = document.getElementById("chkExtraInformation") as HTMLInputElement;
                        const theComparisonType: HTMLInputElement = document.querySelector("input[type='radio'][name='comparison_type']:checked") as HTMLInputElement;
                        const theDisplayType: HTMLInputElement = document.querySelector("input[type='radio'][name='display_type']:checked") as HTMLInputElement;
                        const displayTypeIsJSON: boolean = (theDisplayType.value === "JSON");

                        document.getElementById(tabContainerID).innerHTML = CommonControlsFactory.getLoadingHtml("Working out changes...");
                        const theVersionHistoryDataset: Map<number, IItemVersionHistory> = VersionTabHtmlHelper.WorkOutChanges(theSiteUrl, theFieldsFormatted, theVersionHistory, theFileVersions, (theExtraInformation.checked), (theComparisonType.value === "JSON"));
                        
                        document.getElementById(tabContainerID).innerHTML = CommonControlsFactory.getLoadingHtml("Getting Tabs...");
                        const theByVersionHtml: string = VersionTabHtmlHelper.GetByVersionHistoryTabHtml(theVersionHistoryDataset, displayTypeIsJSON);
                        const theByFieldHtml: string = VersionTabHtmlHelper.GetByFieldHtml(theFieldsFormatted, theVersionHistoryDataset, displayTypeIsJSON);

                        const theVersionIDs: Array<Number> = HelperGenericFunctions.GetKeyArrayFromMap(theVersionHistoryDataset);
                        theVersionIDs.sort(HelperGenericFunctions.DescendingSort);         

                        let theVersionDropList: string = `<select id="{{ID}}">`;
                        theVersionIDs.forEach((currVersionID: number) => {
                          const theVersion = theVersionHistoryDataset.get(currVersionID);
                          theVersionDropList += (`
                            <option value="${currVersionID}">
                              ${theVersion.VersionLabel} (${theVersion.VersionModifiedDate})
                            </option>`
                          );
                        });
                        theVersionDropList += `</select>`;

                        const theCompareByTabHtml: string = (`
                          <div class="${styles.divCompareOptions}">
                            <div class="${styles.divCompareLabel}">
                              Versions to Compare:
                            </div>
                            <div class="${styles.divCompareControl}">
                              ${theVersionDropList.replace("{{ID}}", "ddlVersionA")}
                            </div>
                            <div class="${styles.divCompareControl}">
                              ${theVersionDropList.replace("{{ID}}", "ddlVersionB")}
                            </div>
                            <div class="${styles.divCompareControl}">
                              <input id="btnCompareVersions" type="button" value="Compare" class="${styles.CustomButton}" />
                            </div>
                          </div>
                          <div id="divCompareResults" class="${styles.divCompareResults}">
                          </div>
                        `);                      
                        
                        const theTabController: TabController = new TabController();
                        theTabController.addTab(new TabControl("By Version", theByVersionHtml));
                        theTabController.addTab(new TabControl("By Field", theByFieldHtml));
                        theTabController.addTab(new TabControl("Compare Specific Versions", theCompareByTabHtml));
                        theTabController.addTabHtmlToPage(tabContainerID);

                        const theCompareVersionsButton: HTMLElement = document.getElementById("btnCompareVersions");
                        theCompareVersionsButton.addEventListener("click", (e) => {

                          const leftVersionIDControl: HTMLSelectElement = document.getElementById("ddlVersionA") as HTMLSelectElement;
                          const rightVersionIDControl: HTMLSelectElement = document.getElementById("ddlVersionB") as HTMLSelectElement;

                          if (leftVersionIDControl.value === rightVersionIDControl.value) {
                            document.getElementById("divCompareResults").innerHTML = ("Both versions selected are the same");
                          } else {

                            const leftVersion: IItemVersionHistory = theVersionHistoryDataset.get(Number(leftVersionIDControl.value));
                            const rightVersion: IItemVersionHistory = theVersionHistoryDataset.get(Number(rightVersionIDControl.value));
                            const theCompareHtml: string = VersionTabHtmlHelper.GetCompareVersionsHtml(theFieldsFormatted, leftVersion, rightVersion, displayTypeIsJSON);
                            document.getElementById("divCompareResults").innerHTML = theCompareHtml;
                          }
                        });
                      };
                      

                      if (theVersionHistory.length > 0) {
                        
                        if ((!HelperGenericFunctions.IsNullUndefinedOrEmpty(theVersionHistory[0]["File_x005f_x0020_x005f_Type"])) && (!HelperGenericFunctions.IsNullUndefinedOrEmpty(theVersionHistory[0]["File_x005f_x0020_x005f_Size"]))) {
                          
                          document.getElementById(tabContainerID).innerHTML = CommonControlsFactory.getLoadingHtml("Getting previous versions file sizes...");
                          const theFileRef = (theVersionHistory[0].FileRef);
                          const theFileVersionApiUrl = (theSiteUrl + "/_api/web/GetFileByServerRelativeUrl('" + theFileRef + "')/Versions?$select=Id,Size");
                          
                          SharePointRestApiHelper.RestApiCall(theHttpClient, theFileVersionApiUrl)
                            .then((theFileVersions: IRestApiResponse) => {
                              const theFileVersionHistory: Map<number, number> = new Map<number, number>();
                              theFileVersions.value.forEach((currFileVersion: ISharePointFileVersion) => {
                                theFileVersionHistory.set(currFileVersion.ID, currFileVersion.Size);
                              });

                              document.getElementById(tabContainerID).innerHTML = CommonControlsFactory.getLoadingHtml("Getting current versions file sizes...");
                              const theLatestFileVersionApiUrl = (theSiteUrl + "/_api/web/GetFileByServerRelativeUrl('" + theFileRef + "')");
                              SharePointRestApiHelper.RestApiCall(theHttpClient, theLatestFileVersionApiUrl).then((theLatestVersion: IRestApiResponse) => {
                                const theValue = theLatestVersion.value;
                                theFileVersionHistory.set(theValue[0].UIVersion, theValue[0].Length);
                                
                                OutputResults(theFileVersionHistory);
                              }).catch((error) => {
                                document.getElementById(tabContainerID).innerHTML = "Error getting file versions:<br/>" + error;      
                              });
                            }).catch((error) => {
                              document.getElementById(tabContainerID).innerHTML = "Error getting file versions:<br/>" + error;      
                            });
                        } else {
                          OutputResults();
                        }
                        
                      } else {
                        document.getElementById(tabContainerID).innerHTML = "No versions results";
                      }                    
                    }).catch((error) => {
                      document.getElementById(tabContainerID).innerHTML = "Error retrieving Item Version Data:<br/>" + error;
                    });
                }
              }).catch((error) => {
                document.getElementById(tabContainerID).innerHTML = "Error retrieving Library Data:<br/>" + error;
              });
            } else {
              document.getElementById(tabContainerID).innerHTML = "Please enter a number for the Item ID";
            }
        } else {
          document.getElementById(tabContainerID).innerHTML = "Please enter a list / library root folder path";          
        }
      } else {
        document.getElementById(tabContainerID).innerHTML = "Please enter a SharePoint site url";
      }
    });
  }

  public render(): void {

    this.domElement.innerHTML = (`
      <div id="${this.INPUT_CONTROL_FORM_ID}" class="${styles.divInputForm}">
      </div>
      <div id="${this.RESULTS_TAB_CONTAINER_ID}" class="${styles.divResults}">
      </div>
    `);

    this.renderInputForm();   
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();    
    return super.onInit();
  }


  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [              
              ]
            }
          ]
        }
      ]
    };
  }
}
