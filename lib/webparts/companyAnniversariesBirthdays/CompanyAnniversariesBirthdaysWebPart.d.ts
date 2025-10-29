import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { DisplayMode, FilterMode } from './models/IEmployee';
export interface ICompanyAnniversariesBirthdaysWebPartProps {
    listName: string;
    displayMode: DisplayMode;
    filterMode: FilterMode;
    showImages: boolean;
    showTitle: boolean;
    centerContent: boolean;
    birthdayColor: string;
    anniversaryColor: string;
    certificationColor: string;
}
export default class CompanyAnniversariesBirthdaysWebPart extends BaseClientSideWebPart<ICompanyAnniversariesBirthdaysWebPartProps> {
    private _isDarkTheme;
    private _listDropdownOptions;
    render(): void;
    protected onInit(): Promise<void>;
    private _loadLists;
    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=CompanyAnniversariesBirthdaysWebPart.d.ts.map