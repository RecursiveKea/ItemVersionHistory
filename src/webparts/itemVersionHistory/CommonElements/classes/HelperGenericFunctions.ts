export abstract class HelperGenericFunctions {

    public static GetFileSizeText(theFileSize: number): string {
        
        let theResult: string = "";
        if (theFileSize === 0) {
            theResult = "0 KB";
        } else if (theFileSize >= 1073741824) {
            theResult = ((((theFileSize / 1024) / 1024) / 1024).toFixed(2).toString() + " GB");
        } else if (theFileSize >= 1048576) {
            theResult = (((theFileSize / 1024) / 1024).toFixed(2).toString() + " MB");
        } else {
            theResult = ((theFileSize / 1024).toFixed(2).toString() + " KB");
        }

        return theResult;
    }
    
    public static AscendingSort(a:unknown, b:unknown): number {
        return (a > b) ? 1 : -1;
    }

    public static DescendingSort(a:unknown, b:unknown): number {
        return (a > b) ? -1 : 1;
    }

    public static GetKeyArrayFromMap(theMap: Map<unknown,unknown>): Array<any> {
        const theResult = new Array<any>();
        theMap.forEach((value: unknown, key: unknown) => {
            theResult.push(key);
        });

        return theResult;
    }

    public static IsNullUndefinedOrEmpty(theValue: unknown): boolean {
        return ((theValue === undefined) || (theValue === null) || (theValue === ""));
    }
    
    public static GetTextFromHtmlElement(theID: string): string {
        const theHtmlElement: HTMLInputElement = document.getElementById(theID) as HTMLInputElement;
        const theResult: string = escape(theHtmlElement.value);
        return theResult;
    }    

    //Source: https://stackoverflow.com/questions/6234773/can-i-escape-html-special-chars-in-javascript
    public static EscapeHtml(theHtml: string): string {
        let theResult: string = "";

        if ((theHtml !== undefined) && (theHtml !== null) && (theHtml !== "")) {
            theResult = (
                theHtml.toString()
                    .replace(/&/g, "&amp;")
                    .replace(/</g, "&lt;")
                    .replace(/>/g, "&gt;")
                    .replace(/"/g, "&quot;")
                    .replace(/'/g, "&#039;")
            );
        }
        return theResult;        
    }

    public static GetDateTimeString(theDate: Date): string {

        const theDay: number = theDate.getDate();
        const theMonth: number = theDate.getMonth();
        const theYear: number = theDate.getFullYear();
        const theHours: number = theDate.getHours();
        const theMinutes: number = theDate.getMinutes();
        const theSeconds: number = theDate.getSeconds();

        const am_pm: string = ((theHours >= 12) ? "pm" : "am");
        let theMonthName = "";
        switch (theMonth) {
            case 0: theMonthName = "Jan"; break;
            case 1: theMonthName = "Feb"; break;
            case 2: theMonthName = "Mar"; break;
            case 3: theMonthName = "Apr"; break;
            case 4: theMonthName = "May"; break;
            case 5: theMonthName = "Jun"; break;
            case 6: theMonthName = "Jul"; break;
            case 7: theMonthName = "Aug"; break;
            case 8: theMonthName = "Sep"; break;            
            case 9: theMonthName = "Oct"; break;
            case 10: theMonthName = "Nov"; break;
            case 11: theMonthName = "Dec"; break;
        }

        const theResult: string = (
            theDay.toString() + " "
            + theMonthName + " "
            + theYear.toString() + " " +
            + ((theHours > 13) ? (theHours - 12).toString() : theHours.toString()) + ":"
            + ((theMinutes <= 9) ? "0": "") + theMinutes.toString() + ":"
            + ((theSeconds <= 9) ? "0": "") + theSeconds.toString() + " "
            + am_pm
        );

        return theResult;
    }

}