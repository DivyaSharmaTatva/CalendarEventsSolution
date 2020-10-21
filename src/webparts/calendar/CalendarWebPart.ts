import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CalendarWebPartStrings';
import Calendar from './components/Calendar';
import { ICalendarProps } from './components/ICalendarProps';

export interface ICalendarWebPartProps {
  strWebPartTitle: string;
  strSiteURL: string;
  strListTitle: string;
  intNoOfItems: number;
  strSeeAllURL: string;
}

export default class CalendarWebPart extends BaseClientSideWebPart <ICalendarWebPartProps> {
  /// <summary>CalendarWebPart class.</summary>

  public render(): void {
    /// <summary>Render method.</summary>

    const element: React.ReactElement<ICalendarProps> = React.createElement(
      Calendar,
      {
        strWebPartTitle: this.properties.strWebPartTitle,
        strSiteURL: this.properties.strSiteURL,
        strListTitle: this.properties.strListTitle,
        intNoOfItems: this.properties.intNoOfItems,
        strSeeAllURL: this.properties.strSeeAllURL,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    /// <summary>Dispose method.</summary>

    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    /// <summary>Configure the Property Pane.</summary>

    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('strWebPartTitle', {
                  label: strings.WebPartTitleFieldLabel,
                  value: 'Calendar Events'
                }),
                PropertyPaneTextField('strSiteURL', {
                  label: strings.SiteURLFieldLabel,
                  value: ''
                }),
                PropertyPaneTextField('strListTitle', {
                  label: strings.ListTitleFieldLabel,
                  value: ''
                }),
                PropertyPaneSlider('intNoOfItems', {
                  label: strings.NoOfItemsFieldLabel,
                  min: 1,
                  max: 10
                }),
                PropertyPaneTextField('strSeeAllURL', {
                  label: strings.SeeAllFieldLabel,
                  value: ''
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
