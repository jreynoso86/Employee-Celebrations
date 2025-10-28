import * as React from 'react';
import type { ICompanyAnniversariesBirthdaysProps } from './ICompanyAnniversariesBirthdaysProps';
import { IEmployeeEvent } from '../models/IEmployee';
interface ICompanyAnniversariesBirthdaysState {
    events: IEmployeeEvent[];
    loading: boolean;
    error: string;
    currentIndex: number;
    gridPage: number;
}
export default class CompanyAnniversariesBirthdays extends React.Component<ICompanyAnniversariesBirthdaysProps, ICompanyAnniversariesBirthdaysState> {
    private employeeService;
    private carouselInterval;
    constructor(props: ICompanyAnniversariesBirthdaysProps);
    componentDidMount(): Promise<void>;
    componentWillUnmount(): void;
    componentDidUpdate(prevProps: ICompanyAnniversariesBirthdaysProps): Promise<void>;
    private _loadEvents;
    private _startCarousel;
    private _stopCarousel;
    private _formatDate;
    private _renderEventCard;
    private _renderGridView;
    private _renderListView;
    private _renderCarouselView;
    render(): React.ReactElement<ICompanyAnniversariesBirthdaysProps>;
}
export {};
//# sourceMappingURL=CompanyAnniversariesBirthdays.d.ts.map