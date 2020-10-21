declare interface ICalendarWebPartStrings {
  ///<summary>ICalendarWebPartStrings interface</summary>

  WebPartTitleFieldLabel: string;
  SiteURLFieldLabel: string;
  ListTitleFieldLabel: string;
  NoOfItemsFieldLabel: string;
  SeeAllFieldLabel: string;
}

declare module 'CalendarWebPartStrings' {
  ///<summary>CalendarWebPartStrings module</summary>

  const strings: ICalendarWebPartStrings;
  export = strings;
}
