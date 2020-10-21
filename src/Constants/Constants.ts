import { IIconProps } from 'office-ui-fabric-react';
import { DialogType } from 'office-ui-fabric-react/lib/Dialog';
import * as strings from 'CalendarListingWebPartStrings';
import styles from '../webparts/calendarListing/components/CalendarListing.module.scss';

export const SHAREICON: IIconProps = { iconName: 'Share' };

export const MODELPROPS = {
  isBlocking: false,
};

export const DIALOGCONTENTPROPS = {
  type: DialogType.largeHeader,
  title: strings.StrPopUpHeader,
  className: styles.calendarListing
};

export const MONTHLIST = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

export const DEFAULTURLINEMAIL = "/DispForm.aspx?ID=";

export const REGEXEMAIL = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/;