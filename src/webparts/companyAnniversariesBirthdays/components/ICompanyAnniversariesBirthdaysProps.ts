import { SPHttpClient } from '@microsoft/sp-http';
import { DisplayMode, FilterMode } from '../models/IEmployee';

export interface ICompanyAnniversariesBirthdaysProps {
  listName: string;
  displayMode: DisplayMode;
  filterMode: FilterMode;
  showImages: boolean;
  showTitle: boolean;
  centerContent: boolean;
  birthdayColor: string;
  anniversaryColor: string;
  certificationColor: string;
  isDarkTheme: boolean;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  hasTeamsContext: boolean;
}
