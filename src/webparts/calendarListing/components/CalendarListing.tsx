import * as React from 'react';
import styles from './CalendarListing.module.scss';
import { ICalendarListingProps } from './ICalendarListingProps';
import ICalendarListingStates from './ICalendarListingStates';
import { escape } from '@microsoft/sp-lodash-subset';
import CommonMethods from '../../../Common/CommonMethods';

export default class CalendarListing extends React.Component<ICalendarListingProps, ICalendarListingStates> {

  constructor(props: ICalendarListingProps) {
    super(props);
    
    this.state = {
      ListURL: this.props.strListURL
    };
  }

  public render(): React.ReactElement<ICalendarListingProps> {
    return (
      <div className={ styles.calendarListing }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <p className={ styles.description }>{escape("Calendar Listing WebPart")}</p>
              <p className={ styles.description }>{escape(this.state.ListURL)}</p>
            </div>
          </div>
        </div>
      </div>
    );
  }
  
  public componentDidUpdate(): void{
    ///<summary>React's componentDidMount method</summary>
    try {
      this.setState({
        ListURL: this.props.strListURL
      });  
    } catch (error) {
      console.log("componentDidMount (CalendarListingWebPart.tsx) --> " + error);
    }
  }
}
