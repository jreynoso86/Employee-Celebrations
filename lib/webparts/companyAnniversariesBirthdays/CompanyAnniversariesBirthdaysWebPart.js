var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneDropdown, PropertyPaneToggle, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import CompanyAnniversariesBirthdays from './components/CompanyAnniversariesBirthdays';
import { DisplayMode, FilterMode } from './models/IEmployee';
import { EmployeeService } from './services/EmployeeService';
var CompanyAnniversariesBirthdaysWebPart = /** @class */ (function (_super) {
    __extends(CompanyAnniversariesBirthdaysWebPart, _super);
    function CompanyAnniversariesBirthdaysWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        _this._listDropdownOptions = [];
        return _this;
    }
    CompanyAnniversariesBirthdaysWebPart.prototype.render = function () {
        var element = React.createElement(CompanyAnniversariesBirthdays, {
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
        });
        ReactDom.render(element, this.domElement);
    };
    CompanyAnniversariesBirthdaysWebPart.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, _super.prototype.onInit.call(this)];
                    case 1:
                        _a.sent();
                        // Load available lists
                        return [4 /*yield*/, this._loadLists()];
                    case 2:
                        // Load available lists
                        _a.sent();
                        return [2 /*return*/, Promise.resolve()];
                }
            });
        });
    };
    CompanyAnniversariesBirthdaysWebPart.prototype._loadLists = function () {
        return __awaiter(this, void 0, void 0, function () {
            var employeeService, lists;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        employeeService = new EmployeeService(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl);
                        return [4 /*yield*/, employeeService.getAllLists()];
                    case 1:
                        lists = _a.sent();
                        this._listDropdownOptions = lists.map(function (list) { return ({
                            key: list.title,
                            text: list.title
                        }); });
                        return [2 /*return*/];
                }
            });
        });
    };
    CompanyAnniversariesBirthdaysWebPart.prototype.onThemeChanged = function (currentTheme) {
        if (!currentTheme) {
            return;
        }
        this._isDarkTheme = !!currentTheme.isInverted;
        var semanticColors = currentTheme.semanticColors;
        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--link', semanticColors.link || null);
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
        }
    };
    CompanyAnniversariesBirthdaysWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(CompanyAnniversariesBirthdaysWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    CompanyAnniversariesBirthdaysWebPart.prototype.getPropertyPaneConfiguration = function () {
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
    };
    return CompanyAnniversariesBirthdaysWebPart;
}(BaseClientSideWebPart));
export default CompanyAnniversariesBirthdaysWebPart;
//# sourceMappingURL=CompanyAnniversariesBirthdaysWebPart.js.map