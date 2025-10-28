export interface IEmployee {
    Id: number;
    Title: string;
    HireDate: Date;
    Birthday: Date;
    Certification: string;
    CertificationExpiration: Date;
}
export interface IEmployeeEvent {
    id: number;
    name: string;
    date: Date;
    type: 'birthday' | 'anniversary' | 'certification';
    originalDate: Date;
    yearsCount?: number;
    certificationName?: string;
}
export declare enum DisplayMode {
    Grid = "grid",
    List = "list",
    Carousel = "carousel"
}
export declare enum FilterMode {
    All = "all",
    Today = "today",
    ThisWeek = "thisWeek",
    ThisMonth = "thisMonth",
    NextMonth = "nextMonth"
}
//# sourceMappingURL=IEmployee.d.ts.map