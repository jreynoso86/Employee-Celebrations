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
import styles from './CompanyAnniversariesBirthdays.module.scss';
import { EmployeeService } from '../services/EmployeeService';
import { DisplayMode } from '../models/IEmployee';
import { escape } from '@microsoft/sp-lodash-subset';
var CompanyAnniversariesBirthdays = /** @class */ (function (_super) {
    __extends(CompanyAnniversariesBirthdays, _super);
    function CompanyAnniversariesBirthdays(props) {
        var _this = _super.call(this, props) || this;
        _this.carouselInterval = null;
        _this.state = {
            events: [],
            loading: true,
            error: '',
            currentIndex: 0,
            gridPage: 0
        };
        _this.employeeService = new EmployeeService(props.spHttpClient, props.siteUrl);
        return _this;
    }
    CompanyAnniversariesBirthdays.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._loadEvents()];
                    case 1:
                        _a.sent();
                        // Start carousel if in carousel mode
                        if (this.props.displayMode === DisplayMode.Carousel) {
                            this._startCarousel();
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    CompanyAnniversariesBirthdays.prototype.componentWillUnmount = function () {
        this._stopCarousel();
    };
    CompanyAnniversariesBirthdays.prototype.componentDidUpdate = function (prevProps) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!(prevProps.listName !== this.props.listName ||
                            prevProps.filterMode !== this.props.filterMode)) return [3 /*break*/, 2];
                        return [4 /*yield*/, this._loadEvents()];
                    case 1:
                        _a.sent();
                        _a.label = 2;
                    case 2:
                        // Manage carousel
                        if (this.props.displayMode === DisplayMode.Carousel && prevProps.displayMode !== DisplayMode.Carousel) {
                            this._startCarousel();
                        }
                        else if (this.props.displayMode !== DisplayMode.Carousel && prevProps.displayMode === DisplayMode.Carousel) {
                            this._stopCarousel();
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    CompanyAnniversariesBirthdays.prototype._loadEvents = function () {
        return __awaiter(this, void 0, void 0, function () {
            var employees, events, err_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        this.setState({ loading: true, error: '' });
                        return [4 /*yield*/, this.employeeService.getEmployees(this.props.listName)];
                    case 1:
                        employees = _a.sent();
                        events = this.employeeService.processEmployeeEvents(employees, this.props.filterMode);
                        this.setState({ events: events, loading: false });
                        return [3 /*break*/, 3];
                    case 2:
                        err_1 = _a.sent();
                        console.error('Error loading events:', err_1);
                        this.setState({
                            error: 'Failed to load employee data. Please check your list configuration.',
                            loading: false
                        });
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    CompanyAnniversariesBirthdays.prototype._startCarousel = function () {
        var _this = this;
        this.carouselInterval = window.setInterval(function () {
            _this.setState(function (prevState) { return ({
                currentIndex: (prevState.currentIndex + 1) % Math.max(prevState.events.length, 1)
            }); });
        }, 5000); // Change every 5 seconds
    };
    CompanyAnniversariesBirthdays.prototype._stopCarousel = function () {
        if (this.carouselInterval !== null) {
            window.clearInterval(this.carouselInterval);
            this.carouselInterval = null;
        }
    };
    CompanyAnniversariesBirthdays.prototype._formatDate = function (date) {
        var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        return "".concat(months[date.getMonth()], " ").concat(date.getDate());
    };
    CompanyAnniversariesBirthdays.prototype._renderEventCard = function (event) {
        var icon = '';
        var title = '';
        var colors = '';
        if (event.type === 'birthday') {
            icon = '🎂';
            title = 'Birthday';
            colors = this.props.birthdayColor;
        }
        else if (event.type === 'anniversary') {
            icon = '🎉';
            title = "".concat(event.yearsCount, " Year").concat(event.yearsCount !== 1 ? 's' : '', " Anniversary");
            colors = this.props.anniversaryColor;
        }
        else if (event.type === 'certification') {
            icon = '🏆';
            title = "Certification: ".concat(event.certificationName);
            colors = '#28a745, #20c997'; // Green gradient for certifications
        }
        // Parse gradient colors
        var _a = colors.split(',').map(function (c) { return c.trim(); }), color1 = _a[0], color2 = _a[1];
        var gradientStyle = {
            background: "linear-gradient(135deg, ".concat(color1, " 0%, ").concat(color2, " 100%)")
        };
        return (React.createElement("div", { key: "".concat(event.type, "-").concat(event.id), className: styles.eventCard, style: gradientStyle },
            this.props.showImages && (React.createElement("div", { className: styles.iconContainer },
                React.createElement("span", { className: styles.icon }, icon))),
            React.createElement("div", { className: styles.eventDetails },
                React.createElement("h3", { className: styles.employeeName }, escape(event.name)),
                React.createElement("p", { className: styles.eventType }, title),
                React.createElement("p", { className: styles.eventDate }, event.type === 'certification' ? 'Congratulations!' : this._formatDate(event.date)))));
    };
    CompanyAnniversariesBirthdays.prototype._renderGridView = function () {
        var _this = this;
        var centerContent = this.props.centerContent;
        var _a = this.state, events = _a.events, gridPage = _a.gridPage;
        if (centerContent) {
            // Show 4 items at a time with pagination
            var itemsPerPage = 4;
            var totalPages_1 = Math.ceil(events.length / itemsPerPage);
            var startIndex = gridPage * itemsPerPage;
            var endIndex = startIndex + itemsPerPage;
            var visibleEvents = events.slice(startIndex, endIndex);
            var handlePrevious = function () {
                _this.setState(function (prevState) { return ({
                    gridPage: prevState.gridPage > 0 ? prevState.gridPage - 1 : prevState.gridPage
                }); });
            };
            var handleNext = function () {
                _this.setState(function (prevState) { return ({
                    gridPage: prevState.gridPage < totalPages_1 - 1 ? prevState.gridPage + 1 : prevState.gridPage
                }); });
            };
            return (React.createElement("div", { className: styles.gridViewWithArrows },
                gridPage > 0 && (React.createElement("button", { className: styles.arrowButton, onClick: handlePrevious, "aria-label": "Previous" }, "\u2039")),
                React.createElement("div", { className: styles.gridViewCentered }, visibleEvents.map(function (event) { return _this._renderEventCard(event); })),
                gridPage < totalPages_1 - 1 && (React.createElement("button", { className: styles.arrowButton, onClick: handleNext, "aria-label": "Next" }, "\u203A"))));
        }
        // Default grid view (not centered)
        return (React.createElement("div", { className: styles.gridView }, events.map(function (event) { return _this._renderEventCard(event); })));
    };
    CompanyAnniversariesBirthdays.prototype._renderListView = function () {
        var _this = this;
        return (React.createElement("div", { className: styles.listView }, this.state.events.map(function (event) {
            var icon = '';
            var title = '';
            if (event.type === 'birthday') {
                icon = '🎂';
                title = 'Birthday';
            }
            else if (event.type === 'anniversary') {
                icon = '🎉';
                title = "".concat(event.yearsCount, " Year").concat(event.yearsCount !== 1 ? 's' : '', " Anniversary");
            }
            else if (event.type === 'certification') {
                icon = '🏆';
                title = "Certification: ".concat(event.certificationName);
            }
            return (React.createElement("div", { key: "".concat(event.type, "-").concat(event.id), className: styles.listItem },
                _this.props.showImages && React.createElement("span", { className: styles.listIcon }, icon),
                React.createElement("div", { className: styles.listContent },
                    React.createElement("span", { className: styles.listName }, escape(event.name)),
                    React.createElement("span", { className: styles.listType }, title)),
                React.createElement("span", { className: styles.listDate }, event.type === 'certification' ? 'Congratulations!' : _this._formatDate(event.date))));
        })));
    };
    CompanyAnniversariesBirthdays.prototype._renderCarouselView = function () {
        var _this = this;
        if (this.state.events.length === 0) {
            return React.createElement("div", { className: styles.noEvents }, "No upcoming celebrations");
        }
        var event = this.state.events[this.state.currentIndex];
        return (React.createElement("div", { className: styles.carouselView },
            this._renderEventCard(event),
            React.createElement("div", { className: styles.carouselIndicators }, this.state.events.map(function (_, index) { return (React.createElement("span", { key: index, className: "".concat(styles.indicator, " ").concat(index === _this.state.currentIndex ? styles.active : ''), onClick: function () { return _this.setState({ currentIndex: index }); } })); }))));
    };
    CompanyAnniversariesBirthdays.prototype.render = function () {
        var _a = this.props, displayMode = _a.displayMode, hasTeamsContext = _a.hasTeamsContext, showTitle = _a.showTitle, centerContent = _a.centerContent;
        var _b = this.state, loading = _b.loading, error = _b.error, events = _b.events;
        return (React.createElement("section", { className: "".concat(styles.companyAnniversariesBirthdays, " ").concat(hasTeamsContext ? styles.teams : '', " ").concat(centerContent ? styles.centered : '') },
            showTitle && (React.createElement("div", { className: styles.header },
                React.createElement("h2", { className: styles.title }, "Company Celebrations"))),
            loading && (React.createElement("div", { className: styles.loading }, "Loading celebrations...")),
            error && (React.createElement("div", { className: styles.error }, error)),
            !loading && !error && events.length === 0 && (React.createElement("div", { className: styles.noEvents }, "No upcoming celebrations for the selected period.")),
            !loading && !error && events.length > 0 && (React.createElement(React.Fragment, null,
                displayMode === DisplayMode.Grid && this._renderGridView(),
                displayMode === DisplayMode.List && this._renderListView(),
                displayMode === DisplayMode.Carousel && this._renderCarouselView()))));
    };
    return CompanyAnniversariesBirthdays;
}(React.Component));
export default CompanyAnniversariesBirthdays;
//# sourceMappingURL=CompanyAnniversariesBirthdays.js.map