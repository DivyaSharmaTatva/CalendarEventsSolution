import * as React from 'react';
import styles from './CalendarListing.module.scss';
import { ICalendarListingProps } from './ICalendarListingProps';
import ICalendarListingStates from './ICalendarListingStates';
import { escape } from '@microsoft/sp-lodash-subset';
import CommonMethods from '../../../Common/CommonMethods';
import * as Constants from '../../../Constants/Constants';
import { ICalendarItemDetails } from "../../../Models/ICalendarItemDetails";
import { Web } from 'sp-pnp-js';
import { IconButton, MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { Dialog, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import * as strings from 'CalendarListingWebPartStrings';
import { IDigestCache, DigestCache } from '@microsoft/sp-http';
import * as moment from 'moment';

import * as $ from 'jquery';
require('../../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
require('../../../Style/style.css');

require('../../../../node_modules/bootstrap/dist/js/bootstrap.bundle.min.js');

export default class CalendarListing extends React.Component<ICalendarListingProps, ICalendarListingStates> {
  /// <summary>class component that is used for Listing the calendar</summary>
  /// <param name="ICalendarListingProps">Properties variable</param>
  /// <param name="ICalendarListingStates">States properties</param>

  private methods: CommonMethods = null;
  private objWeb: Web;

  constructor(props: ICalendarListingProps) {
    super(props);
    this.methods = new CommonMethods();

    this.state = {
      ListURL: this.methods.ReturnArgumentString(this.props.strListURL),
      CalendarItems: [],
      ClsShareButton: styles.displayNone,
      HiddenDialog: true,
      MessageBarType: MessageBarType.success,
      MessageHidden: styles.displayNone,
      MessageText: "",
      FromEmailAddress: ""
    };
  }

  public async componentDidMount() {
    ///<summary>React's componentDidMount method</summary>

    try {
      if (this.props.strFromEmailAddress !== null && this.props.strFromEmailAddress !== undefined && this.props.strFromEmailAddress.trim() !== "") {
        this.setState({
          FromEmailAddress: this.methods.ReturnArgumentString(this.props.strFromEmailAddress)
        });
      }
      else {
        this.setState({ FromEmailAddress: "" });
      }

      const TitleCon: Element = this.props.anyObjDom.querySelector('#displayMsg');
      if (this.props.strListURL !== null && this.props.strListURL !== undefined && this.props.strListURL.trim() !== "") {
        this.setState({
          ListURL: this.methods.ReturnArgumentString(this.props.strListURL)
        });
        TitleCon.innerHTML = ``;
        this.RenderItems();
      }
      else {
        TitleCon.innerHTML = strings.MsgAddURL;
        this.setState({ CalendarItems: [] });
      }

      if (Boolean(this.props.strRequiredShare)) {
        this.setState({ ClsShareButton: styles.displayBlock });
      }
      else {
        this.setState({ ClsShareButton: styles.displayNone });
      }
    }
    catch (exception) {
      console.log('componentDidMount (CalendarListing.tsx) --> ' + exception);
    }
  }

  public render(): React.ReactElement<ICalendarListingProps> {
    ///<summary>React's render method</summary>

    try {
      return (
        <div className={styles.calendarListing}>
          <div className="col-md-6 col-lg-4 col-xl-3 mb-4">
            <div className="card p-2 h-100 upcoming-events">
              <div className="card-header view-all-block">
                <img className="icon" src={require('../../../Images/calendar.svg')} alt="Events" />
                {strings.StrUpcomingEvents}
                <a href={this.state.ListURL} title={strings.StrSeeall} className="view-all-btn">{strings.StrSeeall}</a>
              </div>
              <div className="card-body events-list">
                <h3 id="displayMsg"></h3>
                <ul className="list-unstyled">
                  {this.state.CalendarItems.map(currentObj => (
                    <li className="media">
                      <div className={`${"icon"} ${"mr-3"} ${currentObj.Month.toLowerCase()}`}>
                        <span className="month-name">{currentObj.Month}</span>
                        <img src={require('../../../Images/' + currentObj.Month + '.svg')} alt="" />
                        <span className="date">{currentObj.Day}</span>
                      </div>
                      <div className="media-body">
                        <h3>{currentObj.Title}</h3>
                        {currentObj.Description.replace(/<[^>]+>/g, '')}
                        <IconButton iconProps={Constants.SHAREICON} className={`${this.state.ClsShareButton}`} title={strings.StrShare} onClick={() => { this.setState({ HiddenDialog: false, ShareId: Number(currentObj.ID) }); }} />
                      </div>
                    </li>
                  ))}
                </ul>
              </div>
            </div>
          </div>

          <Dialog
            hidden={this.state.HiddenDialog}
            onDismiss={() => this.CloseCancelClick()}
            dialogContentProps={Constants.DIALOGCONTENTPROPS}
            modalProps={Constants.MODELPROPS}>

            <div className={this.state.MessageHidden}>
              <MessageBar className="custom-message" messageBarType={this.state.MessageBarType} isMultiline={true}>{this.state.MessageText.split('\n').map(i => {
                return <div>{i}</div>;
              })}</MessageBar>
            </div>
            <div className="form-group">
              <label>{strings.StrEmailRecipient}<sub>*</sub></label>
              <div className="date-icon">
                <input type="text" placeholder={"abc@xyz.com"} maxLength={200} className="form-control datepicker" value={this.state.StrEmailReciepent} onChange={(event) => { this.setState({ StrEmailReciepent: event.target.value }); }} ></input>
              </div>
            </div>
            <DialogFooter>
              <button type="button" className="btn btn-primary" disabled={this.state.SendDisable} onClick={() => this.ShareEvent()} ><span>{strings.StrShare}</span></button>
              <button type="button" className="btn btn-outline-primary" disabled={this.state.SendDisable} onClick={() => this.CloseCancelClick()}><span>{strings.StrCancel}</span></button>
            </DialogFooter>
          </Dialog>
        </div>
      );
    }
    catch (exception) {
      console.log('render (CalendarListing.tsx) --> ' + exception);
    }
  }

  public componentDidUpdate(prevProps): void {
    ///<summary>React's componentDidUpdate method</summary>
    /// <param name="prevProps">Previous Properties variable</param>

    try {
      if (prevProps.strListURL !== this.props.strListURL) {
        const TitleCon: Element = this.props.anyObjDom.querySelector('#displayMsg');
        if (this.props.strListURL !== null && this.props.strListURL !== undefined && this.props.strListURL.trim() !== "") {
          this.setState({
            ListURL: this.methods.ReturnArgumentString(this.props.strListURL)
          });
          TitleCon.innerHTML = ``;
          this.RenderItems();
        }
        else {
          TitleCon.innerHTML = strings.MsgAddURL;
          this.setState({ CalendarItems: [] });
        }
      }

      if (prevProps.intNoOfItems !== this.props.intNoOfItems) {
        if (this.props.intNoOfItems !== null && this.props.intNoOfItems !== undefined) {
          this.RenderItems();
        }
      }

      if (prevProps.strRequiredShare !== this.props.strRequiredShare) {
        if (Boolean(this.props.strRequiredShare)) {
          this.setState({ ClsShareButton: styles.displayBlock });
        }
        else {
          this.setState({ ClsShareButton: styles.displayNone });
        }
      }

      if (prevProps.strFromEmailAddress !== this.props.strFromEmailAddress) {
        if (this.props.strFromEmailAddress !== null && this.props.strFromEmailAddress !== undefined && this.props.strFromEmailAddress.trim() !== "") {
          this.setState({
            FromEmailAddress: this.methods.ReturnArgumentString(this.props.strFromEmailAddress)
          });
        }
        else {
          this.setState({ FromEmailAddress: "" });
        }
      }
    }
    catch (exception) {
      console.log('componentDidUpdate (CalendarListing.tsx) --> ' + exception);
    }
  }

  private RenderItems() {
    /// <summary>Render Calendar Items</summary>

    let NoOfItems = 5;
    try {
      const TitleCon: Element = this.props.anyObjDom.querySelector('#displayMsg');
      if (!isNaN(Number(this.props.intNoOfItems)) && Number(this.props.intNoOfItems) > 0) {
        NoOfItems = Number(this.props.intNoOfItems);
      }

      this.methods.GetCalendarListItems(this.props.strListURL, NoOfItems).then((lstCalendarItem: ICalendarItemDetails[]) => {
        if (lstCalendarItem !== null && lstCalendarItem.length > 0) {
          for (let loopAllItmes = 0; loopAllItmes < lstCalendarItem.length; loopAllItmes++) {
            if (lstCalendarItem[loopAllItmes].Description === null) {
              lstCalendarItem[loopAllItmes].Description = "";
            }
  
            let currentEventDate = new Date(String(lstCalendarItem[loopAllItmes].EventDate));
            lstCalendarItem[loopAllItmes].Month = Constants.MONTHLIST[currentEventDate.getMonth()];
            lstCalendarItem[loopAllItmes].Day = currentEventDate.getDate();
          }

          this.setState({ CalendarItems: lstCalendarItem });
          TitleCon.innerHTML = ``;
        }
        else {
          TitleCon.innerHTML = strings.StrItemNotFound;
          this.setState({ CalendarItems: [] });
        }
      });
    }
    catch (exception) {
      console.log('RenderItems (CalendarListing.tsx) --> ' + exception);
    }
  }

  private CloseCancelClick() {
    /// <summary>close dialog and clear state</summary>

    this.setState({ MessageHidden: styles.displayNone, MessageBarType: MessageBarType.success, MessageText: "", StrEmailReciepent: "", HiddenDialog: true, SendDisable: false });
  }

  public getUserDetailsByEmail(email: string): Promise<string> {
    /// <summary>close dialog and clear state</summary>
    /// <param name="email">Email</param>

    this.objWeb = new Web(this.props.anyPageContext.web.absoluteUrl);
    return this.objWeb.ensureUser(email).then(result => {
      return result.data.Title;
    }).catch((exception) => {
      console.log('getUserDetailsByEmail (CalendarListing.tsx) -->' + exception);
      return Promise.reject(JSON.stringify(exception));
    });
  }

  private ShareEvent() {
    /// <summary>Share Event</summary>

    try {
      let errMsgText = "";
      this.setState({ SendDisable: true });

      if (this.state.FromEmailAddress.length <= 0) {
        errMsgText = strings.ErrFromEmailAddress;
        this.setState({ MessageHidden: styles.displayBlock, SendDisable: false, MessageBarType: MessageBarType.error, MessageText: errMsgText });
        let $this = this;
        setTimeout(() => {
          $this.setState({ MessageHidden: styles.displayNone, MessageBarType: MessageBarType.success, MessageText: "" });
        }, 20000);
      }
      else {
        if (!this.methods.ValidateString(this.state.StrEmailReciepent)) {
          errMsgText = "'" + strings.StrEmailRecipient + "' " + strings.ErrGeneralMsg;
        }
        else if (!Constants.REGEXEMAIL.test(this.state.StrEmailReciepent)) {
          errMsgText = strings.ErrEmailFormat;
        }

        if (errMsgText.length > 0) {
          this.setState({ MessageHidden: styles.displayBlock, SendDisable: false, MessageBarType: MessageBarType.error, MessageText: errMsgText });
          let $this = this;
          setTimeout(() => {
            $this.setState({ MessageHidden: styles.displayNone, MessageBarType: MessageBarType.success, MessageText: "" });
          }, 20000);
        }
        else {
          const digestCache: IDigestCache = this.props.anyContext.serviceScope.consume(DigestCache.serviceKey);
          digestCache.fetchDigest(this.props.anyPageContext.web.serverRelativeUrl).then((digest: string): void => {
            let sitetemplateurl = window.origin + "/_api/SP.Utilities.Utility.SendEmail";
            let pageURL = this.props.strListURL + Constants.DEFAULTURLINEMAIL + this.state.ShareId;
            let EmailBody = "";
            let EmailSubject = "";

            this.getUserDetailsByEmail(this.state.StrEmailReciepent).then((userName) => {
              this.methods.GetCalendarListItemByID(this.props.strListURL, this.state.ShareId).then((lstCalendarItemDetails: ICalendarItemDetails) => {
                if (lstCalendarItemDetails) {
                  EmailSubject = 'Invitation to event: ' + lstCalendarItemDetails.Title;

                  var tempadate = new Date(moment.utc(lstCalendarItemDetails.EventDate).format());
                  moment.locale(this.props.anyPageContext.legacyPageContext["currentCultureName"]);
                  let convertedEventDate = moment(tempadate, "DD-MM-yyyy hh:mm A").format('L');
                  let Description = lstCalendarItemDetails.Description == null ? "" : lstCalendarItemDetails.Description.replace(/<[^>]+>/g, '');
                  let location = lstCalendarItemDetails.Location == null ? "" : lstCalendarItemDetails.Location;
                  let Category = lstCalendarItemDetails.Category == null ? "" : lstCalendarItemDetails.Category;

                  EmailBody = 'Hello ' + userName + ',<br/><br/><b>Title:</b> ' + lstCalendarItemDetails.Title + '<br/><b>Description:</b> ' + Description + '<br/><b>Event Date:</b> ' + convertedEventDate + '<br/><b>Location:</b> ' + location + '<br/><b>Category:</b> ' + Category + '<br/><br/>If you have access on this event then visit it in SharePoint by clicking on link <a href="' + pageURL + '">here</a>.<br/><br/>Thanks.';

                  let $this = this;

                  $.ajax({
                    contentType: 'application/json',
                    url: sitetemplateurl,
                    type: "POST",
                    data: JSON.stringify({
                      'properties': {
                        '__metadata': { 'type': 'SP.Utilities.EmailProperties' },
                        'From': this.state.FromEmailAddress,
                        'To': { 'results': [this.state.StrEmailReciepent] },
                        'Body': EmailBody,
                        'Subject': EmailSubject
                      }
                    }),
                    headers: {
                      "Accept": "application/json;odata=verbose",
                      "content-type": "application/json;odata=verbose",
                      "X-RequestDigest": digest,
                    },
                    success: (data) => {
                      this.setState({ StrEmailReciepent: "", HiddenDialog: true, SendDisable: false });
                    },
                    error: (err) => {
                      console.log('ShareEvent error (CalendarListing.tsx)--> ' + JSON.stringify(err));
                    }
                  });
                }
              });
            });
          });
        }
      }
    }
    catch (exception) {
      console.log('ShareEvent (CalendarListing.tsx) --> ' + exception);
    }
  }
}
