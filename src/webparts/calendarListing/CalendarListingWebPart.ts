import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CalendarListingWebPartStrings';
import CalendarListing from './components/CalendarListing';
import { ICalendarListingProps } from './components/ICalendarListingProps';

export interface ICalendarListingWebPartProps {
  strListURL: string;
  intNoOfItems: number;  
}

export default class CalendarListingWebPart extends BaseClientSideWebPart <ICalendarListingWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICalendarListingProps> = React.createElement(
      CalendarListing,
      {
        strListURL: this.properties.strListURL,
        intNoOfItems: this.properties.intNoOfItems,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
