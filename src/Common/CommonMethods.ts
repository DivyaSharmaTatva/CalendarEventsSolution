import { Web, ListEnsureResult } from 'sp-pnp-js';
import { ICalendarItems } from "../Models/ICalendarItems";

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
}