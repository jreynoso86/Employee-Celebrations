import { SPHttpClient } from '@microsoft/sp-http';
import { IEmployee, IEmployeeEvent, FilterMode } from '../models/IEmployee';
export declare class EmployeeService {
    private spHttpClient;
    private siteUrl;
    constructor(spHttpClient: SPHttpClient, siteUrl: string);
    /**
     * Create the Employee Celebrations list if it doesn't exist
     */
    ensureListExists(): Promise<boolean>;
    /**
     * Get all employees from the specified list
     */
    getEmployees(listName: string): Promise<IEmployee[]>;
    /**
     * Get all SharePoint lists from the current site
     */
    getAllLists(): Promise<Array<{
        title: string;
        id: string;
    }>>;
    /**
     * Process employees and calculate upcoming events
     */
    processEmployeeEvents(employees: IEmployee[], filterMode?: FilterMode): IEmployeeEvent[];
    /**
     * Calculate the next occurrence of a date (birthday or anniversary)
     */
    private calculateNextOccurrence;
    /**
     * Calculate years of service
     */
    private calculateYearsOfService;
    /**
     * Filter events based on the selected filter mode
     */
    private filterEvents;
    /**
     * Check if two dates are the same day
     */
    private isSameDay;
}
//# sourceMappingURL=EmployeeService.d.ts.map