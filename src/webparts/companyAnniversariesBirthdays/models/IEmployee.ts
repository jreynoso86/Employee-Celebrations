export interface IEmployee {
  Id: number;
  Title: string; // Employee Name
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
  yearsCount?: number; // For anniversaries
  certificationName?: string; // For certifications
}

export enum DisplayMode {
  Grid = 'grid',
  List = 'list',
  Carousel = 'carousel'
}

export enum FilterMode {
  All = 'all',
  Today = 'today',
  ThisWeek = 'thisWeek',
  ThisMonth = 'thisMonth',
  NextMonth = 'nextMonth'
}
