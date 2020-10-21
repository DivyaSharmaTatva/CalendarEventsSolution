import { MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
import { ICalendarItemDetails } from "../../../Models/ICalendarItemDetails";

export default interface ICalendarListingStates{
    ///<summary>ICalendarListingStates interface</summary>
    ListURL : string;

    CalendarItems:ICalendarItemDetails[];
    ClsShareButton:string;
    HiddenDialog:boolean;
    MessageHidden:string;
    MessageBarType: MessageBarType;
    MessageText:string;
    StrEmailReciepent?:string;
    SendDisable?:boolean;
    ShareId?:number;
    FromEmailAddress?:string;
}