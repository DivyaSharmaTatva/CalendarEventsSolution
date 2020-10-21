export interface ICalendarItemDetails {
    ID: string;
    Title: string;
    Description: string;
    EventDate: Date;
    Day?:number;
    Month?:string;
    Location?:string;
    Category?:string;
}