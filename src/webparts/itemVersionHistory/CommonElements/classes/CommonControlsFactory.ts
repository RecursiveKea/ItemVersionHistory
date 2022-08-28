import GenericControlStyles from '../stylesheets/CommonControlsFactoryStyle.module.scss'
import { HelperGenericFunctions } from '../classes/HelperGenericFunctions';
 
export abstract class CommonControlsFactory {

    public static wrapLabelAndInputsInTemplateHtml(theLabel: string, theInputsHtmls: string): string {
        return (`
            <div class="${GenericControlStyles.divGenericControlWrapper}">
                <div class="${GenericControlStyles.divGenericControlLabelWrapper}">
                    <label>${theLabel}</label>
                </div>
                <div class="${GenericControlStyles.divGenericInputTextControlWrapper}">
                    ${theInputsHtmls}
                </div>
            </div>
        `);
    }

    public static getInputControlWithLabelWithoutWrappingDivs(theID: string, theLabel: string, theControlType: string, inputBeforeLabel: boolean = false, theAdditionalInputAttributesHtml: string = ""): string {
        let theResult: string = "";
        if (inputBeforeLabel) {
            theResult = (`        
                <input id="${theID}" type="${theControlType}" ${theAdditionalInputAttributesHtml} />
                <label for="${theID}">
                    ${theLabel}
                </label>                        
            `);
        } else {
            theResult = (`        
                <label for="${theID}">
                    ${theLabel}
                </label>        
                <input id="${theID}" type="${theControlType}" ${theAdditionalInputAttributesHtml} />
            `);
        }

        return theResult;
    }    

    public static getInputTextControlHtml(theID: string, theLabel: string, theTextBoxWidthPx?: number): string {
        let theTextBoxWidthStyle = "";
        if (!HelperGenericFunctions.IsNullUndefinedOrEmpty(theTextBoxWidthPx)) {
            theTextBoxWidthStyle = (` style="width: ${theTextBoxWidthPx.toString()}px" `);
        }

        const theResult = (`
            <div class="${GenericControlStyles.divGenericControlWrapper}">
                <div class="${GenericControlStyles.divGenericControlLabelWrapper}">
                    <label for="${theID}">
                        ${theLabel}
                    </label>
                </div>
                <div class="${GenericControlStyles.divGenericInputTextControlWrapper}">
                    <input id="${theID}" type="Text" ${theTextBoxWidthStyle} />
                </div>
            </div>
        `);

        return theResult;
    }

    public static getInputButtonControlHtml(theID: string, theLabel: string, theButtonClass: string = ""): string {
        const theResult = (`
            <div class="${GenericControlStyles.divGenericControlWrapper}">
                <div class="${GenericControlStyles.divGenericInputButtonControlWrapper}">
                    <input id="${theID}" type="button" value="${theLabel}" class="${theButtonClass}" />
                </div>
            </div>            
        `);

        return theResult;
    }


    public static getInputCheckboxControlHtml(theID: string, theLabel: string): string {
        const theResult = (`
            <div class="${GenericControlStyles.divGenericControlWrapper}">                
                <div class="${GenericControlStyles.divGenericControlLabelWrapper}">
                    <label for="${theID}">
                        ${theLabel}
                    </label>
                </div>
                <div class="${GenericControlStyles.divGenericInputCheckboxControlWrapper}">
                    <input id="${theID}" type="Text" />
                </div>
            </div>            
        `);

        return theResult;
    }    


    public static getLoadingHtml(theText: string): string {
        return (`
            <div>
                <div class="${GenericControlStyles.ldsRing}">
                    <div></div>
                    <div></div>
                    <div></div>
                    <div></div>
                </div>
                <br/>
                ${theText}
            </div>            
        `);
      }
}