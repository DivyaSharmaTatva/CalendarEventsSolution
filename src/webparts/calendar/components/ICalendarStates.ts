import { ICalendarItems } from "../../../Models/ICalendarItems";

export default interface ICalendarStates{
    ///<summary>ICalendarStates interface</summary>
    
    isListExist: boolean;
    arrCalendarItems : ICalendarItems[];
}