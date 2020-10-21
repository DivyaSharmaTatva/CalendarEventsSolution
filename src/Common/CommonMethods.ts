import { Web, ListEnsureResult } from 'sp-pnp-js';
import { ICalendarItems } from "../Models/ICalendarItems";
import { ICalendarItemDetails } from "../Models/ICalendarItemDetails";

export default class CommonMethods {

    private objWeb: Web;

    public EnsureListExists(strSiteURL: string, strListTitle: string) : Promise<boolean> {
        /// <summary>Method checks whether the list exists on the provided Site URL.</summary>
        /// <param name="strSiteURL">Site URL</param>
        /// <param name="strListTitle">List Title</param>

        let isListExist: boolean = false;

        try {
            this.objWeb = new Web(strSiteURL);
            
            return this.objWeb.lists.ensure(strListTitle).then((list: ListEnsureResult) => {                                
                if(list.created ==  false) {
                    isListExist = true;
                }
                else {
                    this.objWeb.lists.getByTitle(strListTitle).delete();
                }

                return isListExist;
            });
        }
        catch (exception) {
            console.error('EnsureListExists (CommonMethods.ts) --> ' + exception);       
            return Promise.resolve(isListExist);
        }
    }

    public GetCalendarEvents = async (strSiteURL: string, strListTitle: string, intNoOfEvents: number) : Promise<ICalendarItems[]> => {
        /// <summary>Fetch the Calendar Events based on provided arguments.</summary>
        /// <param name="strSiteURL">Site URL</param>
        /// <param name="strListTitle">List Title</param>
        /// <param name="intNoOfEvents">No. of events to fetch</param>
        /// <returns name="ICalendarItems">Returns Calendar Events</returns>

        try {
            this.objWeb = new Web(strSiteURL);
            let intTopItemsCount =  intNoOfEvents != null ? intNoOfEvents : 5;

            return this.objWeb.lists.getByTitle(strListTitle).items
                .select('ID', 'Title', 'Description', 'EventDate', 'Location')
                .top(intTopItemsCount)
                .orderBy('EventDate')
                .filter(`EventDate ge datetime'${new Date().toISOString()}'`)
                .get().then((lstCalendarItem: ICalendarItems[]) => {
                    return lstCalendarItem;
                })
                .catch((exception) => {
                    console.error('GetCalendarListItems 1st catch (CommonMethods.ts) --> ' + exception);
                    return null;
                });
        }
        catch (exception) {
            console.error('GetCalendarListItems 2nd catch (CommonMethods.ts) --> ' + exception);
            return null;
        }
    } 

    public ReturnArgumentString(strListURL: string): string {
        return strListURL;
    }
    
    public GetCalendarListItems(strListURL: string, intNoOfItems: number): Promise<ICalendarItemDetails[]> {
        /// <summary>Get calendar list items</summary>
        /// <param name="strListURL">strListURL</param>
        /// <param name="intNoOfItems">intNoOfItems</param>

        try {
            let currentURL = new URL(strListURL);
            let listURL = currentURL.pathname;

            let currentCalendarAbsoluteURL = strListURL.substr(0, strListURL.lastIndexOf('/Lists/'));
            this.objWeb = new Web(currentCalendarAbsoluteURL);

            let objCurrentDateTime = new Date().toISOString();

            return this.objWeb.getList(listURL).items
                .select('ID', 'Title', 'Description', 'EventDate')
                .top(intNoOfItems)
                .orderBy('EventDate')
                .filter(`EventDate ge datetime'${objCurrentDateTime}'`)
                .get().then((lstCalendarItem: ICalendarItems[]) => {
                    return lstCalendarItem;
                }).catch((exception) => {
                    console.log('GetCalendarListItems 1st catch (CommonMethods.ts) --> ' + exception);
                    return null;
                });
        }
        catch (exception) {
            console.log('GetCalendarListItems 2nd catch (CommonMethods.ts) --> ' + exception);
        }
    }

    public ValidateString(value) {
        /// <summary>validate string type value</summary>
        /// <param name="value">value</param>

        if (value === undefined || value === "") {
            return false;
        }
        else if (value.trim() == "") {
            return false;
        }
        else {
            return true;
        }
    }

    public GetCalendarListItemByID(strListURL: string, eventId: number): Promise<ICalendarItemDetails> {
        /// <summary>Get calendar list item by Id</summary>
        /// <param name="strListURL">strListURL</param>
        /// <param name="eventId">eventId</param>

        try {
            let currentURL = new URL(strListURL);
            let listURL = currentURL.pathname;

            let currentCalendarAbsoluteURL = strListURL.substr(0, strListURL.lastIndexOf('/Lists/'));
            this.objWeb = new Web(currentCalendarAbsoluteURL);

            return this.objWeb.getList(listURL).items
                .getById(eventId)
                .select('ID', 'Title', 'Description', 'EventDate', 'Location', 'Category')
                .get()
                .then((lstCalendarItemDetails: ICalendarItemDetails) => {
                    return lstCalendarItemDetails;
                }).catch((exception) => {
                    console.log('GetCalendarListItemByID 1st catch (CommonMethods.ts) --> ' + exception);
                    return null;
                });
        }
        catch (exception) {
            console.log('GetCalendarListItemByID 2nd catch (CommonMethods.ts) --> ' + exception);
        }
    }
}