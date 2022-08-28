import TabStyle from '../stylesheets/TabStyle.module.scss'

export class TabControl {

    public HeadingHtml: string;
    public ContentHtml: string;

    constructor(theHeadingHtml: string, contentHtml: string) {
        this.HeadingHtml = theHeadingHtml;
        this.ContentHtml = contentHtml;
    }
}


export class TabController {

    private _tabs: Array<TabControl>;  

    constructor() {
        this._tabs = new Array<TabControl>();
    }

    addTab(theTab: TabControl) {
        this._tabs.push(theTab);
    }

    addTabHtmlToPage(theContainerDivID: string): void {
        
        if (this._tabs.length === 0) {
            throw new Error("No tabs to display");
        } else {
            
            let theTabHeaders: string = `<div class="${TabStyle.divAllTabHeadingWrapper}">`;
            let theTabContent: string = `<div class="${TabStyle.divAllTabContentWrapper}">`;

            let theCurrentTabIndex: number = 0;
            this._tabs.forEach((currTab: TabControl) => {
                let tabStateClass: string = `${TabStyle.InActiveTab }`;
                if (theCurrentTabIndex === 0) {
                    tabStateClass = `${TabStyle.ActiveTab }`;
                }

                theTabHeaders += (`
                    <div class="${TabStyle.divTabHeadingWrapper} ${tabStateClass}" data-tabIndex="${theCurrentTabIndex}">
                        ${currTab.HeadingHtml}
                    </div>
                `);

                theTabContent += (`
                    <div class="${TabStyle.divTabContentWrapper} ${tabStateClass}" data-tabIndex="${theCurrentTabIndex}">
                        ${currTab.ContentHtml}
                    </div>
                `);

                theCurrentTabIndex++;
            });

            theTabHeaders += "</div>"
            theTabContent += "</div>"

            const theTabHtml: string = (`
                <div class="${TabStyle.divTabControlWrapper}">
                    ${theTabHeaders}
                    ${theTabContent}
                </div>
            `);

            document.getElementById(theContainerDivID).innerHTML = theTabHtml;

            const divTabHeadingsForWiringUpClickEvent = document.querySelectorAll(`.${ TabStyle.divTabHeadingWrapper }`);            
            divTabHeadingsForWiringUpClickEvent.forEach((currTabHeadingForClickEvent: HTMLElement) => {
                currTabHeadingForClickEvent.addEventListener("click", function(e) {                                        

                    const theCurrentTabHeadingAndContent = document.querySelectorAll(`.${TabStyle.ActiveTab }`);
                    theCurrentTabHeadingAndContent.forEach((currTab) => {
                        currTab.className = currTab.className.replace(`${TabStyle.ActiveTab}`, `${TabStyle.InActiveTab}`);
                    });
                    

                    const newCurrentTabIndex = this.getAttribute("data-tabIndex");
                    const theNewTabHeadingAndContent = document.querySelectorAll(`.${TabStyle.InActiveTab}[data-tabIndex="${newCurrentTabIndex}"]`);
                    theNewTabHeadingAndContent.forEach((currTab) => {
                        currTab.className = currTab.className.replace(`${TabStyle.InActiveTab}`, `${TabStyle.ActiveTab}`);
                    });
                });
            });
        }        
    }
}