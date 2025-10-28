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
import { SPHttpClient } from '@microsoft/sp-http';
import { FilterMode } from '../models/IEmployee';
var EmployeeService = /** @class */ (function () {
    function EmployeeService(spHttpClient, siteUrl) {
        this.spHttpClient = spHttpClient;
        this.siteUrl = siteUrl;
    }
    /**
     * Create the Employee Celebrations list if it doesn't exist
     */
    EmployeeService.prototype.ensureListExists = function () {
        return __awaiter(this, void 0, void 0, function () {
            var listName, checkResponse, createListResponse, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 7, , 8]);
                        listName = 'Employee Celebrations';
                        return [4 /*yield*/, this.spHttpClient.get("".concat(this.siteUrl, "/_api/web/lists/getbytitle('").concat(listName, "')"), SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=nometadata',
                                    'odata-version': ''
                                }
                            })];
                    case 1:
                        checkResponse = _a.sent();
                        if (checkResponse.ok) {
                            return [2 /*return*/, true]; // List already exists
                        }
                        return [4 /*yield*/, this.spHttpClient.post("".concat(this.siteUrl, "/_api/web/lists"), SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=nometadata',
                                    'Content-type': 'application/json;odata=nometadata',
                                    'odata-version': ''
                                },
                                body: JSON.stringify({
                                    '__metadata': { 'type': 'SP.List' },
                                    'AllowContentTypes': true,
                                    'BaseTemplate': 100,
                                    'ContentTypesEnabled': true,
                                    'Description': 'List to track employee birthdays and work anniversaries',
                                    'Title': listName
                                })
                            })];
                    case 2:
                        createListResponse = _a.sent();
                        if (!createListResponse.ok) {
                            console.error('Failed to create list');
                            return [2 /*return*/, false];
                        }
                        // Add HireDate field
                        return [4 /*yield*/, this.spHttpClient.post("".concat(this.siteUrl, "/_api/web/lists/getbytitle('").concat(listName, "')/fields"), SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=nometadata',
                                    'Content-type': 'application/json;odata=nometadata',
                                    'odata-version': ''
                                },
                                body: JSON.stringify({
                                    '__metadata': { 'type': 'SP.FieldDateTime' },
                                    'FieldTypeKind': 4,
                                    'Title': 'HireDate',
                                    'DisplayFormat': 1
                                })
                            })];
                    case 3:
                        // Add HireDate field
                        _a.sent();
                        // Add Birthday field
                        return [4 /*yield*/, this.spHttpClient.post("".concat(this.siteUrl, "/_api/web/lists/getbytitle('").concat(listName, "')/fields"), SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=nometadata',
                                    'Content-type': 'application/json;odata=nometadata',
                                    'odata-version': ''
                                },
                                body: JSON.stringify({
                                    '__metadata': { 'type': 'SP.FieldDateTime' },
                                    'FieldTypeKind': 4,
                                    'Title': 'Birthday',
                                    'DisplayFormat': 1
                                })
                            })];
                    case 4:
                        // Add Birthday field
                        _a.sent();
                        // Add Certification field
                        return [4 /*yield*/, this.spHttpClient.post("".concat(this.siteUrl, "/_api/web/lists/getbytitle('").concat(listName, "')/fields"), SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=nometadata',
                                    'Content-type': 'application/json;odata=nometadata',
                                    'odata-version': ''
                                },
                                body: JSON.stringify({
                                    '__metadata': { 'type': 'SP.Field' },
                                    'FieldTypeKind': 2,
                                    'Title': 'Certification'
                                })
                            })];
                    case 5:
                        // Add Certification field
                        _a.sent();
                        // Add CertificationExpiration field
                        return [4 /*yield*/, this.spHttpClient.post("".concat(this.siteUrl, "/_api/web/lists/getbytitle('").concat(listName, "')/fields"), SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=nometadata',
                                    'Content-type': 'application/json;odata=nometadata',
                                    'odata-version': ''
                                },
                                body: JSON.stringify({
                                    '__metadata': { 'type': 'SP.FieldDateTime' },
                                    'FieldTypeKind': 4,
                                    'Title': 'CertificationExpiration',
                                    'DisplayFormat': 1
                                })
                            })];
                    case 6:
                        // Add CertificationExpiration field
                        _a.sent();
                        return [2 /*return*/, true];
                    case 7:
                        error_1 = _a.sent();
                        console.error('Error ensuring list exists:', error_1);
                        return [2 /*return*/, false];
                    case 8: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get all employees from the specified list
     */
    EmployeeService.prototype.getEmployees = function (listName) {
        return __awaiter(this, void 0, void 0, function () {
            var response, data, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.spHttpClient.get("".concat(this.siteUrl, "/_api/web/lists/getbytitle('").concat(listName, "')/items?$select=Id,Title,HireDate,Birthday,Certification,CertificationExpiration"), SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=nometadata',
                                    'odata-version': ''
                                }
                            })];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("Failed to fetch employees: ".concat(response.statusText));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        return [2 /*return*/, data.value.map(function (item) { return ({
                                Id: item.Id,
                                Title: item.Title,
                                HireDate: item.HireDate ? new Date(item.HireDate) : null,
                                Birthday: item.Birthday ? new Date(item.Birthday) : null,
                                Certification: item.Certification || null,
                                CertificationExpiration: item.CertificationExpiration ? new Date(item.CertificationExpiration) : null
                            }); })];
                    case 3:
                        error_2 = _a.sent();
                        console.error('Error fetching employees:', error_2);
                        return [2 /*return*/, []];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get all SharePoint lists from the current site
     */
    EmployeeService.prototype.getAllLists = function () {
        return __awaiter(this, void 0, void 0, function () {
            var response, data, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.spHttpClient.get("".concat(this.siteUrl, "/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100&$select=Title,Id"), SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=nometadata',
                                    'odata-version': ''
                                }
                            })];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("Failed to fetch lists: ".concat(response.statusText));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        return [2 /*return*/, data.value.map(function (list) { return ({
                                title: list.Title,
                                id: list.Id
                            }); })];
                    case 3:
                        error_3 = _a.sent();
                        console.error('Error fetching lists:', error_3);
                        return [2 /*return*/, []];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Process employees and calculate upcoming events
     */
    EmployeeService.prototype.processEmployeeEvents = function (employees, filterMode) {
        var _this = this;
        if (filterMode === void 0) { filterMode = FilterMode.All; }
        var events = [];
        var today = new Date();
        today.setHours(0, 0, 0, 0); // Reset time to start of day for accurate comparison
        var currentYear = today.getFullYear();
        employees.forEach(function (employee) {
            // Process Birthday
            if (employee.Birthday) {
                var nextBirthday = _this.calculateNextOccurrence(employee.Birthday, currentYear);
                events.push({
                    id: employee.Id,
                    name: employee.Title,
                    date: nextBirthday,
                    type: 'birthday',
                    originalDate: employee.Birthday
                });
            }
            // Process Anniversary
            if (employee.HireDate) {
                var nextAnniversary = _this.calculateNextOccurrence(employee.HireDate, currentYear);
                var yearsOfService = _this.calculateYearsOfService(employee.HireDate, nextAnniversary);
                events.push({
                    id: employee.Id,
                    name: employee.Title,
                    date: nextAnniversary,
                    type: 'anniversary',
                    originalDate: employee.HireDate,
                    yearsCount: yearsOfService
                });
            }
            // Process Certification
            // Only show certifications if today's date is less than or equal to expiration date
            if (employee.Certification && employee.CertificationExpiration) {
                var expirationDate = new Date(employee.CertificationExpiration.toString());
                expirationDate.setHours(0, 0, 0, 0); // Reset time for accurate comparison
                // Only include if certification has not expired
                if (today.getTime() <= expirationDate.getTime()) {
                    var currentDate = new Date();
                    currentDate.setHours(0, 0, 0, 0);
                    events.push({
                        id: employee.Id,
                        name: employee.Title,
                        date: currentDate, // Show certification immediately
                        type: 'certification',
                        originalDate: employee.CertificationExpiration,
                        certificationName: employee.Certification
                    });
                }
            }
        });
        // Filter events based on filter mode
        var filteredEvents = this.filterEvents(events, filterMode, today);
        // Sort by date
        return filteredEvents.sort(function (a, b) { return a.date.getTime() - b.date.getTime(); });
    };
    /**
     * Calculate the next occurrence of a date (birthday or anniversary)
     */
    EmployeeService.prototype.calculateNextOccurrence = function (originalDate, currentYear) {
        var today = new Date();
        var month = originalDate.getMonth();
        var day = originalDate.getDate();
        // Try current year
        var nextDate = new Date(currentYear, month, day);
        // If the date has already passed this year, use next year
        if (nextDate.getTime() < today.getTime()) {
            nextDate = new Date(currentYear + 1, month, day);
        }
        return nextDate;
    };
    /**
     * Calculate years of service
     */
    EmployeeService.prototype.calculateYearsOfService = function (hireDate, anniversaryDate) {
        return anniversaryDate.getFullYear() - hireDate.getFullYear();
    };
    /**
     * Filter events based on the selected filter mode
     */
    EmployeeService.prototype.filterEvents = function (events, filterMode, today) {
        var _this = this;
        switch (filterMode) {
            case FilterMode.Today:
                return events.filter(function (event) { return _this.isSameDay(event.date, today); });
            case FilterMode.ThisWeek: {
                var endOfWeek_1 = new Date(today.getTime());
                endOfWeek_1.setDate(today.getDate() + (7 - today.getDay()));
                return events.filter(function (event) { return event.date.getTime() >= today.getTime() && event.date.getTime() <= endOfWeek_1.getTime(); });
            }
            case FilterMode.ThisMonth:
                return events.filter(function (event) {
                    return event.date.getMonth() === today.getMonth() &&
                        event.date.getFullYear() === today.getFullYear();
                });
            case FilterMode.NextMonth: {
                var nextMonth_1 = new Date(today.getTime());
                nextMonth_1.setMonth(today.getMonth() + 1);
                return events.filter(function (event) {
                    return event.date.getMonth() === nextMonth_1.getMonth() &&
                        event.date.getFullYear() === nextMonth_1.getFullYear();
                });
            }
            case FilterMode.All:
            default:
                return events;
        }
    };
    /**
     * Check if two dates are the same day
     */
    EmployeeService.prototype.isSameDay = function (date1, date2) {
        return date1.getDate() === date2.getDate() &&
            date1.getMonth() === date2.getMonth() &&
            date1.getFullYear() === date2.getFullYear();
    };
    return EmployeeService;
}());
export { EmployeeService };
//# sourceMappingURL=EmployeeService.js.map