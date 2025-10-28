import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import CompanyAnniversariesBirthdays from './components/CompanyAnniversariesBirthdays';
import { ICompanyAnniversariesBirthdaysProps } from './components/ICompanyAnniversariesBirthdaysProps';
import { DisplayMode, FilterMode } from './models/IEmployee';
import { EmployeeService } from './services/EmployeeService';

export interface ICompanyAnniversariesBirthdaysWebPartProps {
  listName: string;
  displayMode: DisplayMode;
  filterMode: FilterMode;
  showImages: boolean;
  showTitle: boolean;
  centerContent: boolean;
  birthdayColor: string;
  anniversaryColor: string;
}

export default class CompanyAnniversariesBirthdaysWebPart extends BaseClientSideWebPart<ICompanyAnniversariesBirthdaysWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _listDropdownOptions: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    const element: React.ReactElement<ICompanyAnniversariesBirthdaysProps> = React.createElement(
      CompanyAnniversariesBirthdays,
      {
        listName: this.properties.listName || 'Employee Celebrations',
        displayMode: this.properties.displayMode || DisplayMode.Grid,
        filterMode: this.properties.filterMode || FilterMode.ThisMonth,
        showImages: this.properties.showImages !== false,
        showTitle: this.properties.showTitle !== false,
        centerContent: this.properties.centerContent === true,
        birthdayColor: this.properties.birthdayColor || '#f093fb,#f5576c',
        anniversaryColor: this.properties.anniversaryColor || '#4facfe,#00f2fe',
        isDarkTheme: this._isDarkTheme,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        hasTeamsContext: !!this.context.sdks.microsoftTeams
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    // Load available lists
    await this._loadLists();

    return Promise.resolve();
  }

  private async _loadLists(): Promise<void> {
    const employeeService = new EmployeeService(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl
    );

    const lists = await employeeService.getAllLists();
    this._listDropdownOptions = lists.map(list => ({
      key: list.title,
      text: list.title
    }));
  }




  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure the Company Anniversaries and Birthdays web part'
          },
          groups: [
            {
              groupName: 'Data Source',
              groupFields: [
                PropertyPaneDropdown('listName', {
                  label: 'Select Employee List',
                  options: this._listDropdownOptions,
                  selectedKey: this.properties.listName || 'Employee Celebrations'
                })
              ]
            },
            {
              groupName: 'Display Settings',
              groupFields: [
                PropertyPaneDropdown('displayMode', {
                  label: 'Display Mode',
                  options: [
                    { key: DisplayMode.Grid, text: 'Grid View' },
                    { key: DisplayMode.List, text: 'List View' },
                    { key: DisplayMode.Carousel, text: 'Carousel View' }
                  ],
                  selectedKey: this.properties.displayMode || DisplayMode.Grid
                }),
                PropertyPaneDropdown('filterMode', {
                  label: 'Show Events',
                  options: [
                    { key: FilterMode.Today, text: 'Today Only' },
                    { key: FilterMode.ThisWeek, text: 'This Week' },
                    { key: FilterMode.ThisMonth, text: 'This Month' },
                    { key: FilterMode.NextMonth, text: 'Next Month' },
                    { key: FilterMode.All, text: 'All Upcoming' }
                  ],
                  selectedKey: this.properties.filterMode || FilterMode.ThisMonth
                }),
                PropertyPaneToggle('showImages', {
                  label: 'Show Images',
                  checked: this.properties.showImages !== false,
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneToggle('showTitle', {
                  label: 'Show Title',
                  checked: this.properties.showTitle !== false,
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneToggle('centerContent', {
                  label: 'Center Content',
                  checked: this.properties.centerContent === true,
                  onText: 'On',
                  offText: 'Off'
                })
              ]
            },
            {
              groupName: 'Colors',
              groupFields: [
                PropertyPaneTextField('birthdayColor', {
                  label: 'Birthday Colors (gradient: color1,color2)',
                  description: 'Enter two colors separated by comma for gradient',
                  placeholder: '#f093fb,#f5576c'
                }),
                PropertyPaneTextField('anniversaryColor', {
                  label: 'Anniversary Colors (gradient: color1,color2)',
                  description: 'Enter two colors separated by comma for gradient',
                  placeholder: '#4facfe,#00f2fe'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
