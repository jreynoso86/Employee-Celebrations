import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IEmployee, IEmployeeEvent, FilterMode } from '../models/IEmployee';

export class EmployeeService {
  private spHttpClient: SPHttpClient;
  private siteUrl: string;

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this.spHttpClient = spHttpClient;
    this.siteUrl = siteUrl;
  }

  /**
   * Create the Employee Celebrations list if it doesn't exist
   */
  public async ensureListExists(): Promise<boolean> {
    try {
      const listName = 'Employee Celebrations';

      // Check if list exists
      const checkResponse = await this.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (checkResponse.ok) {
        return true; // List already exists
      }

      // Create the list
      const createListResponse = await this.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists`,
        SPHttpClient.configurations.v1,
        {
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
        }
      );

      if (!createListResponse.ok) {
        console.error('Failed to create list');
        return false;
      }

      // Add HireDate field
      await this.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
        SPHttpClient.configurations.v1,
        {
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
        }
      );

      // Add Birthday field
      await this.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
        SPHttpClient.configurations.v1,
        {
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
        }
      );

      // Add Certification field
      await this.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
        SPHttpClient.configurations.v1,
        {
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
        }
      );

      // Add CertificationExpiration field
      await this.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
        SPHttpClient.configurations.v1,
        {
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
        }
      );

      return true;
    } catch (error) {
      console.error('Error ensuring list exists:', error);
      return false;
    }
  }

  /**
   * Get all employees from the specified list
   */
  public async getEmployees(listName: string): Promise<IEmployee[]> {
    try {
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,HireDate,Birthday,Certification,CertificationExpiration`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch employees: ${response.statusText}`);
      }

      const data = await response.json();
      return data.value.map((item: { Id: number; Title: string; HireDate: string; Birthday: string; Certification: string; CertificationExpiration: string }) => ({
        Id: item.Id,
        Title: item.Title,
        HireDate: item.HireDate ? new Date(item.HireDate) : null,
        Birthday: item.Birthday ? new Date(item.Birthday) : null,
        Certification: item.Certification || null,
        CertificationExpiration: item.CertificationExpiration ? new Date(item.CertificationExpiration) : null
      }));
    } catch (error) {
      console.error('Error fetching employees:', error);
      return [];
    }
  }

  /**
   * Get all SharePoint lists from the current site
   */
  public async getAllLists(): Promise<Array<{ title: string; id: string }>> {
    try {
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100&$select=Title,Id`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch lists: ${response.statusText}`);
      }

      const data = await response.json();
      return data.value.map((list: { Title: string; Id: string }) => ({
        title: list.Title,
        id: list.Id
      }));
    } catch (error) {
      console.error('Error fetching lists:', error);
      return [];
    }
  }

  /**
   * Process employees and calculate upcoming events
   */
  public processEmployeeEvents(employees: IEmployee[], filterMode: FilterMode = FilterMode.All): IEmployeeEvent[] {
    const events: IEmployeeEvent[] = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Reset time to start of day for accurate comparison
    const currentYear = today.getFullYear();

    employees.forEach(employee => {
      // Process Birthday
      if (employee.Birthday) {
        const nextBirthday = this.calculateNextOccurrence(employee.Birthday, currentYear);
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
        const nextAnniversary = this.calculateNextOccurrence(employee.HireDate, currentYear);
        const yearsOfService = this.calculateYearsOfService(employee.HireDate, nextAnniversary);
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
        const expirationDate = new Date(employee.CertificationExpiration.toString());
        expirationDate.setHours(0, 0, 0, 0); // Reset time for accurate comparison

        // Only include if certification has not expired
        if (today.getTime() <= expirationDate.getTime()) {
          const currentDate = new Date();
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
    const filteredEvents = this.filterEvents(events, filterMode, today);

    // Sort by date
    return filteredEvents.sort((a, b) => a.date.getTime() - b.date.getTime());
  }

  /**
   * Calculate the next occurrence of a date (birthday or anniversary)
   */
  private calculateNextOccurrence(originalDate: Date, currentYear: number): Date {
    const today = new Date();
    const month = originalDate.getMonth();
    const day = originalDate.getDate();

    // Try current year
    let nextDate = new Date(currentYear, month, day);

    // If the date has already passed this year, use next year
    if (nextDate.getTime() < today.getTime()) {
      nextDate = new Date(currentYear + 1, month, day);
    }

    return nextDate;
  }

  /**
   * Calculate years of service
   */
  private calculateYearsOfService(hireDate: Date, anniversaryDate: Date): number {
    return anniversaryDate.getFullYear() - hireDate.getFullYear();
  }

  /**
   * Filter events based on the selected filter mode
   */
  private filterEvents(events: IEmployeeEvent[], filterMode: FilterMode, today: Date): IEmployeeEvent[] {
    switch (filterMode) {
      case FilterMode.Today:
        return events.filter(event => this.isSameDay(event.date, today));

      case FilterMode.ThisWeek: {
        const endOfWeek = new Date(today.getTime());
        endOfWeek.setDate(today.getDate() + (7 - today.getDay()));
        return events.filter(event => event.date.getTime() >= today.getTime() && event.date.getTime() <= endOfWeek.getTime());
      }

      case FilterMode.ThisMonth:
        return events.filter(event =>
          event.date.getMonth() === today.getMonth() &&
          event.date.getFullYear() === today.getFullYear()
        );

      case FilterMode.NextMonth: {
        const nextMonth = new Date(today.getTime());
        nextMonth.setMonth(today.getMonth() + 1);
        return events.filter(event =>
          event.date.getMonth() === nextMonth.getMonth() &&
          event.date.getFullYear() === nextMonth.getFullYear()
        );
      }

      case FilterMode.All:
      default:
        return events;
    }
  }

  /**
   * Check if two dates are the same day
   */
  private isSameDay(date1: Date, date2: Date): boolean {
    return date1.getDate() === date2.getDate() &&
           date1.getMonth() === date2.getMonth() &&
           date1.getFullYear() === date2.getFullYear();
  }
}
