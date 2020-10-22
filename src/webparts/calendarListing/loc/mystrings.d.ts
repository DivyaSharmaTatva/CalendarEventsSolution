declare interface ICalendarListingWebPartStrings {
  ListURLFieldLabel: string;
  NoOfItemsFieldLabel: string;

  RequiredShare:string;
  ToggleOnText: string;
  ToggleOffText: string;
  StrEmailRecipient :string;
  StrPopUpHeader:string;
  ErrGeneralMsg:string;
  StrUpcomingEvents:string;
  StrSeeall:string;
  StrShare:string;
  StrCancel:string;
  MsgAddURL:string;
  ErrEmailFormat:string;
  FromEmailAddressFieldLabel:string;
  ErrFromEmailAddress:string;
  StrItemNotFound:string;
}

declare module 'CalendarListingWebPartStrings' {
  const strings: ICalendarListingWebPartStrings;
  export = strings;
}
