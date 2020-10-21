import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CalendarListingWebPartStrings';
import CalendarListing from './components/CalendarListing';
import { ICalendarListingProps } from './components/ICalendarListingProps';

export default class CalendarListingWebPart extends BaseClientSideWebPart <ICalendarListingProps> {

  public render(): void {
    const element: React.ReactElement<ICalendarListingProps> = React.createElement(
      CalendarListing,
      {
        strListURL: this.properties.strListURL,
        intNoOfItems: this.properties.intNoOfItems,
        strRequiredShare: this.properties.strRequiredShare,
        anyPageContext: this.context.pageContext,
        anyContext:this.context,
        anyObjDom:this.domElement,
        strFromEmailAddress:this.properties.strFromEmailAddress
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('strListURL', {
                  label: strings.ListURLFieldLabel,
                  value: ''
                }),
                PropertyPaneTextField('intNoOfItems', {
                  label: strings.NoOfItemsFieldLabel,
                  value: '' //Number textbox
                }),
                PropertyPaneTextField('strFromEmailAddress', {
                  label: strings.FromEmailAddressFieldLabel,
                  value: ''
                }),
                PropertyPaneToggle('strRequiredShare', {
                  label: strings.RequiredShare,
                  onText: '' + strings.ToggleOnText + '',
                  offText: '' + strings.ToggleOffText + ''
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
