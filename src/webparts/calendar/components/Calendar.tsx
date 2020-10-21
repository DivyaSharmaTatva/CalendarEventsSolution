import * as React from 'react';
import styles from './Calendar.module.scss';
import { ICalendarProps } from './ICalendarProps';
import ICalendarStates from './ICalendarStates';
import { escape } from '@microsoft/sp-lodash-subset';
import CommonMethods from '../../../Common/CommonMethods';
import { resultItem } from 'office-ui-fabric-react/lib/components/ExtendedPicker/PeoplePicker/ExtendedPeoplePicker.scss';
import moment from 'moment';

require('../../../Style/bootstrap.min.css');
require('../../../Style/style.css');

const strWPTitle : string = "Calendar Events";
const strNoRecords : string = "No records found.";
const strWebPartPropertiesBindMSg: string = "Bind 'Site URL' and 'List Title' webpart properties to show the records.";
const strListNotExist : string = "List does not exist.";
const strNoLocation: string = "-";

export default class Calendar extends React.Component<ICalendarProps, ICalendarStates> {
  
  private objCommonMethods: CommonMethods = null;
  private strDefaultSeeAll: string = "#";

  constructor(props: ICalendarProps) {
    super(props);
    this.objCommonMethods = new CommonMethods();
    
    this.state = {      
      isListExist: false,
      arrCalendarItems: []
    };    
  }
  
  public render(): React.ReactElement<ICalendarProps> {
    /// <summary>Render the DOM elements to show the Calendar Events.</summary>
    
    return (
      <div className={styles.calendar}>
        <main className={styles["main-content"]}>
			    <div className="container-fluid">
				    <div className="row">
              <div className="col-xl-9 mb-4">
                <div className={"card p-2 h-100 upcoming-events " + styles.fixHeight }>
                  <div className="card-header view-all-block">
                    <img className="icon" src={require('../../../Images/calendar.svg')} alt="" />
                    {this.props.strWebPartTitle != undefined && this.props.strWebPartTitle.trim().length > 0 ? this.props.strWebPartTitle : strWPTitle}
                    <a data-interception="off" href={(this.props.strSeeAllURL != null || this.props.strSeeAllURL != undefined) ? this.props.strSeeAllURL : this.strDefaultSeeAll} title="See all" className="view-all-btn">See all</a>
                  </div>
                  <div className="card-body customScroll events-list">
                    <div className="container-fluid">
									    <div className= "list-unstyled row">
                        { this.bindWebPart() }
                      </div>
                    </div>
                  </div>                
                </div>
              </div>        
            </div>
          </div>
        </main>
      </div>
    );
  }
  
  public componentDidMount = () => {
    /// <summary>Method to bind the Calendar Events while mounting.</summary>
    
    try {
      if ((this.props.strSiteURL != null || this.props.strSiteURL != undefined)
      && (this.props.strListTitle != null || this.props.strListTitle != undefined)) {
        this.fetchCalendarEvents();
      }
    } catch (error) {
      console.error("componentDidMount (Calendar.tsx) --> " + error);
    }    
  }

  public componentDidUpdate = (prevProps: ICalendarProps) => {
    /// <summary>Method to bind the Calendar Events when there is any update.</summary>
    /// <param name="prevProps">Previous Properties</param>

    try {   
      let isUpdateRequire: boolean = false;

      if (prevProps.strSiteURL !== this.props.strSiteURL) {
        if (this.props.strSiteURL !== null || this.props.strSiteURL !== undefined || this.props.strSiteURL !== "") {
          isUpdateRequire = true;
        }
      }
      
      if (prevProps.strListTitle !== this.props.strListTitle) {
        if (this.props.strListTitle !== null || this.props.strListTitle !== undefined || this.props.strListTitle !== "") {
          isUpdateRequire = true;
        }
      }
      
      if (prevProps.intNoOfItems !== this.props.intNoOfItems) {
        if (this.props.intNoOfItems !== null || this.props.intNoOfItems !== undefined) {
          isUpdateRequire = true;
        }
      }

      if (prevProps.strSeeAllURL !== this.props.strSeeAllURL) {
        if (this.props.strSeeAllURL !== null || this.props.strSeeAllURL !== undefined || this.props.strSeeAllURL !== "") {
          isUpdateRequire = true;
        }
      }
      
      if (isUpdateRequire) {
        this.fetchCalendarEvents();
      }      
    } catch (error) {
      console.error("componentDidUpdate (Calendar.tsx) --> " + error);
    }
  }

  private fetchCalendarEvents = () => {
    /// <summary>Fetch Calendar Events based on provided properties.</summary>
    
    try {
      this.objCommonMethods.EnsureListExists(this.props.strSiteURL, this.props.strListTitle).then((isExists) => {
        if(isExists ==  true) {
          this.objCommonMethods.GetCalendarEvents(this.props.strSiteURL, this.props.strListTitle, this.props.intNoOfItems).then((objCalendarItems) => {
            this.setState({     
              isListExist: true,         
              arrCalendarItems: objCalendarItems
            });  
          });
        }
         else {
          this.setState({        
            isListExist: false,      
            arrCalendarItems: []
          });
        }
      });
    } catch (error) {
        console.error("fetchCalendarEvents (Calendar.tsx) --> " + error);
    }    
  }

  private bindWebPart = () => {
    /// <summary>Bind the WebPart Content.</summary>

    try {
      if (this.props.strSiteURL != null && this.props.strSiteURL != undefined && this.props.strSiteURL != "" && this.props.strListTitle != null && this.props.strListTitle != undefined && this.props.strListTitle != "") {
        if(this.state.isListExist) {
          if (this.state.arrCalendarItems != null && this.state.arrCalendarItems.length > 0 ) {
            return this.bindCalendarEvents();
          } 
          else {          
            return this.bindMessage(strNoRecords); 
          }
        }
        else {
          return this.bindMessage(strListNotExist);
        }
      }
      else {
        return this.bindMessage(strWebPartPropertiesBindMSg);
      }      
    }
    catch(error){
      console.error("bindWebPart (Calendar.tsx) --> " + error);
    }    
  }

  private bindCalendarEvents = () => {
    ///<summary>Bind Calendar Events to DOM.</summary>

    try {
      return(
        this.state.arrCalendarItems.map((item)=>
          <div className='col-md-4'>
            <div className="media">
              <div className={"icon mr-3 " + (moment(item.EventDate).format("MMM")).toString().toLowerCase()}>
                <span className="month-name">{moment(item.EventDate).format("MMM")}</span>
                <img src={require('../../../Images/' + moment(item.EventDate).format("MMM") + '.svg')} alt="" />
                <span className="date">{moment(item.EventDate).date()}</span>	
              </div>
              <div className="media-body">
                <h3>{item.Title}</h3>
                {item.Description != null && item.Description.length > 0 ? item.Description.replace(/<[^>]+>/g, '') : ''}
                <p className={styles.location + " d-flex align-items-center"}><img className={styles.icon} src={require('../../../Images/location.svg')} alt=""/> {item.Location != null && item.Location.length > 0 ? item.Location : strNoLocation}</p>
              </div>
          </div>
        </div>        
        )
      );
    } catch (error) {
      console.error("bindCalendarEvents (Calendar.tsx) --> " + error);  
    }
  }

  private bindMessage = (strMessage: string) => {
    ///<summary>Method to bind message in the WebPart.</summary>
    /// <param name="strMessage">Message string</param>

    try {
      return (
        <div className={styles.message}>
          <h3 className="mb-1">{strMessage}</h3>
        </div>
      );
    } catch (error) {
      console.error("bindMessage (Calendar.tsx) --> " + error);
    }
  }
}
