import { SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import { IRestApiResponse } from '../interfaces/IRestApiResponse';
import { HelperGenericFunctions } from '../classes/HelperGenericFunctions';


export abstract class SharePointRestApiHelper {

    public static RestApiCall(theHttpClient: SPHttpClient, theRestApiUrl: string): Promise<IRestApiResponse> {
        let theResult: IRestApiResponse = { value: new Array() };
        
        return theHttpClient.get(theRestApiUrl, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {                
                if (!response.ok) {
                    throw new Error("Error retrieving record set: " + response.status.toString());
                }

                return response.json();    
            })
            .then((responseJSON: IRestApiResponse) => {
                if (responseJSON.value !== undefined) {                
                    responseJSON.value.forEach((currElement: any) => {
                        theResult.value.push(currElement);
                    });

                    const theNextLink = responseJSON['@odata.nextLink'];
                    if ((theNextLink !== undefined) && (theNextLink !== "")) {
                        return this.RestApiCall(theHttpClient, theNextLink).then((recursiveResponse: IRestApiResponse) => {
                            recursiveResponse.value.forEach((currElement: any) => {
                                theResult.value.push(currElement);
                            });

                            return theResult;
                        });
                    }
                } else {
                    theResult.value.push(responseJSON);
                }
  
                return theResult;
            });
    }

    //Used for: "Created_x0020_By" to "Created_x005f_x0020_x005f_By" and names starting with "_" to "OData__x005f_"
    public static DecodeFieldNameFromRestResponse(theFieldName: string): string {            
        let theResult = theFieldName;

        if (theResult.substring(0, 1) === "_") {
            theResult = ("OData__x005f" + theResult);
        }

        if (theResult.indexOf("OData__x005f_") !== -1) {            
            theResult = ("OData__x005f_" + theResult.replace("OData__x005f_", "").replace(/_/g, "_x005f_"));
        } else {
            theResult = (theResult.replace(/_/g, "_x005f_"));
        }

        return theResult;
    }

    public static GetListFieldValueFromRestResult(theField: string, theFieldType: string, theValue: any, theValueToSetIfBlank?: string, showExtraContent: boolean = false): string {
        let theResult: string = "";        
        theField = this.DecodeFieldNameFromRestResponse(theField);


        if (theValue[theField] === undefined) {
            throw new Error("The field '" + theField + "' was not found in the obnject passed in");
        } else if ((theValue[theField] === "") || (theValue[theField] === null)) {
            theResult = ((theValueToSetIfBlank !== null) ? theValueToSetIfBlank : "");
        } else {
            switch (theFieldType) {
                case "DateTime": {
                    const theDate: Date = new Date(theValue[theField]);
                    theResult = HelperGenericFunctions.GetDateTimeString(theDate);
                }
                break;

                case "URL":
                    theResult = (
                        theValue[theField].Url +
                       (showExtraContent ? " (Description: " + theValue[theField].Description + ")": "")
                    );                
                    break;

                case "Lookup": {
                    if (theValue[theField].LookupValue === undefined) {
                        //Some lookups (eg ContentVersion) don't refer to a list ...and can be a date (eg Created_x0020_Date)
                        if ((theValue[theField][10] === "T") && (theValue[theField][19] === "Z")) {
                            const theDate: Date = new Date(theValue[theField]);
                            theResult = HelperGenericFunctions.GetDateTimeString(theDate);
                        } else if (theField === "File_X0020_Size") {
                            theResult = HelperGenericFunctions.GetFileSizeText(Number(theValue[theField]));
                        } else {
                            theResult = theValue[theField];
                        }                       
                    } else {
                        theResult = (
                            theValue[theField].LookupValue +
                            (showExtraContent ? " (ID: " + theValue[theField].LookupId + ")" : "")
                        );
                    }
                }      
                break;

                case "LookupMulti": {
                    if (theValue[theField].length === 0) {
                        theResult = ((theValueToSetIfBlank !== null) ? theValueToSetIfBlank : "");
                    } else {
                        theValue[theField].forEach((currLookupVal: any) => {                       
                            theResult += (
                                currLookupVal.LookupValue +
                                (showExtraContent ? " (ID: " + currLookupVal.LookupId + ")": "") +
                                "; "
                            );
                        });
    
                        theResult = theResult.substring(0, theResult.length - 2);
                    }                    
                }
                break;

                case "TaxonomyFieldType":
                    theResult = (
                        theValue[theField].Label +
                        (showExtraContent ? " (TermGuid: " + theValue[theField].TermGuid + ")": "")
                    );
                    break;

                case "TaxonomyFieldTypeMulti": {
                    if (theValue[theField].length === 0) {
                        theResult = ((theValueToSetIfBlank !== null) ? theValueToSetIfBlank : "");
                    } else {
                        theValue[theField].forEach((currTerm: any) => {                
                            theResult = (
                                currTerm.Label +
                                (showExtraContent ? " (TermGuid: " + currTerm.TermGuid + ")" : "") +
                                "; "
                            );
                        });

                        theResult = theResult.substring(0, theResult.length - 2);
                    }
                }
                break;

                case "User":
                    theResult = (
                        theValue[theField].LookupValue +
                        (showExtraContent ? " (Email: " + theValue[theField].Email + ")": "")
                    );                
                    break;

                case "UserMulti": {
                    if (theValue[theField].length === 0) {
                        theResult = ((theValueToSetIfBlank !== null) ? theValueToSetIfBlank : "");
                    } else {
                        theValue[theField].forEach((currUser: any) => {
                            theResult += (
                                currUser.LookupValue +
                                (showExtraContent ? " (Email: " + currUser.Email + ")" : "") + 
                                "; "
                            );
                        });

                        theResult = theResult.substring(0, theResult.length - 2);
                    }
                }
                break;

                case "ContentTypeId":
                    theResult = theValue[theField].StringValue;
                    break;

                default:
                    //ModStat, File, Text, Note, Guid, Boolean
                    theResult = theValue[theField];                
                    break;
            }
        }

        theResult = HelperGenericFunctions.EscapeHtml(theResult);
        return theResult;
    }
}