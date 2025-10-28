import * as React from 'react';
import styles from './CompanyAnniversariesBirthdays.module.scss';
import type { ICompanyAnniversariesBirthdaysProps } from './ICompanyAnniversariesBirthdaysProps';
import { EmployeeService } from '../services/EmployeeService';
import { IEmployeeEvent, DisplayMode } from '../models/IEmployee';
import { escape } from '@microsoft/sp-lodash-subset';

interface ICompanyAnniversariesBirthdaysState {
  events: IEmployeeEvent[];
  loading: boolean;
  error: string;
  currentIndex: number;
  gridPage: number;
}

export default class CompanyAnniversariesBirthdays extends React.Component<ICompanyAnniversariesBirthdaysProps, ICompanyAnniversariesBirthdaysState> {
  private employeeService: EmployeeService;
  private carouselInterval: number | null = null;

  constructor(props: ICompanyAnniversariesBirthdaysProps) {
    super(props);

    this.state = {
      events: [],
      loading: true,
      error: '',
      currentIndex: 0,
      gridPage: 0
    };

    this.employeeService = new EmployeeService(
      props.spHttpClient,
      props.siteUrl
    );
  }

  public async componentDidMount(): Promise<void> {
    await this._loadEvents();

    // Start carousel if in carousel mode
    if (this.props.displayMode === DisplayMode.Carousel) {
      this._startCarousel();
    }
  }

  public componentWillUnmount(): void {
    this._stopCarousel();
  }

  public async componentDidUpdate(prevProps: ICompanyAnniversariesBirthdaysProps): Promise<void> {
    // Reload data if props changed
    if (prevProps.listName !== this.props.listName ||
        prevProps.filterMode !== this.props.filterMode) {
      await this._loadEvents();
    }

    // Manage carousel
    if (this.props.displayMode === DisplayMode.Carousel && prevProps.displayMode !== DisplayMode.Carousel) {
      this._startCarousel();
    } else if (this.props.displayMode !== DisplayMode.Carousel && prevProps.displayMode === DisplayMode.Carousel) {
      this._stopCarousel();
    }
  }

  private async _loadEvents(): Promise<void> {
    try {
      this.setState({ loading: true, error: '' });

      const employees = await this.employeeService.getEmployees(this.props.listName);
      const events = this.employeeService.processEmployeeEvents(employees, this.props.filterMode);

      this.setState({ events, loading: false });
    } catch (err) {
      console.error('Error loading events:', err);
      this.setState({
        error: 'Failed to load employee data. Please check your list configuration.',
        loading: false
      });
    }
  }

  private _startCarousel(): void {
    this.carouselInterval = window.setInterval(() => {
      this.setState(prevState => ({
        currentIndex: (prevState.currentIndex + 1) % Math.max(prevState.events.length, 1)
      }));
    }, 5000); // Change every 5 seconds
  }

  private _stopCarousel(): void {
    if (this.carouselInterval !== null) {
      window.clearInterval(this.carouselInterval);
      this.carouselInterval = null;
    }
  }

  private _formatDate(date: Date): string {
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return `${months[date.getMonth()]} ${date.getDate()}`;
  }

  private _renderEventCard(event: IEmployeeEvent): React.ReactElement {
    let icon: string = '';
    let title: string = '';
    let colors: string = '';

    if (event.type === 'birthday') {
      icon = '🎂';
      title = 'Birthday';
      colors = this.props.birthdayColor;
    } else if (event.type === 'anniversary') {
      icon = '🎉';
      title = `${event.yearsCount} Year${event.yearsCount !== 1 ? 's' : ''} Anniversary`;
      colors = this.props.anniversaryColor;
    } else if (event.type === 'certification') {
      icon = '🏆';
      title = `Certification: ${event.certificationName}`;
      colors = '#28a745, #20c997'; // Green gradient for certifications
    }

    // Parse gradient colors
    const [color1, color2] = colors.split(',').map(c => c.trim());
    const gradientStyle = {
      background: `linear-gradient(135deg, ${color1} 0%, ${color2} 100%)`
    };

    return (
      <div key={`${event.type}-${event.id}`} className={styles.eventCard} style={gradientStyle}>
        {this.props.showImages && (
          <div className={styles.iconContainer}>
            <span className={styles.icon}>{icon}</span>
          </div>
        )}
        <div className={styles.eventDetails}>
          <h3 className={styles.employeeName}>{escape(event.name)}</h3>
          <p className={styles.eventType}>{title}</p>
          <p className={styles.eventDate}>{event.type === 'certification' ? 'Congratulations!' : this._formatDate(event.date)}</p>
        </div>
      </div>
    );
  }

  private _renderGridView(): React.ReactElement {
    const { centerContent } = this.props;
    const { events, gridPage } = this.state;

    if (centerContent) {
      // Show 4 items at a time with pagination
      const itemsPerPage = 4;
      const totalPages = Math.ceil(events.length / itemsPerPage);
      const startIndex = gridPage * itemsPerPage;
      const endIndex = startIndex + itemsPerPage;
      const visibleEvents = events.slice(startIndex, endIndex);

      const handlePrevious = (): void => {
        this.setState(prevState => ({
          gridPage: prevState.gridPage > 0 ? prevState.gridPage - 1 : prevState.gridPage
        }));
      };

      const handleNext = (): void => {
        this.setState(prevState => ({
          gridPage: prevState.gridPage < totalPages - 1 ? prevState.gridPage + 1 : prevState.gridPage
        }));
      };

      return (
        <div className={styles.gridViewWithArrows}>
          {gridPage > 0 && (
            <button className={styles.arrowButton} onClick={handlePrevious} aria-label="Previous">
              ‹
            </button>
          )}
          <div className={styles.gridViewCentered}>
            {visibleEvents.map(event => this._renderEventCard(event))}
          </div>
          {gridPage < totalPages - 1 && (
            <button className={styles.arrowButton} onClick={handleNext} aria-label="Next">
              ›
            </button>
          )}
        </div>
      );
    }

    // Default grid view (not centered)
    return (
      <div className={styles.gridView}>
        {events.map(event => this._renderEventCard(event))}
      </div>
    );
  }

  private _renderListView(): React.ReactElement {
    return (
      <div className={styles.listView}>
        {this.state.events.map(event => {
          let icon: string = '';
          let title: string = '';

          if (event.type === 'birthday') {
            icon = '🎂';
            title = 'Birthday';
          } else if (event.type === 'anniversary') {
            icon = '🎉';
            title = `${event.yearsCount} Year${event.yearsCount !== 1 ? 's' : ''} Anniversary`;
          } else if (event.type === 'certification') {
            icon = '🏆';
            title = `Certification: ${event.certificationName}`;
          }

          return (
            <div key={`${event.type}-${event.id}`} className={styles.listItem}>
              {this.props.showImages && <span className={styles.listIcon}>{icon}</span>}
              <div className={styles.listContent}>
                <span className={styles.listName}>{escape(event.name)}</span>
                <span className={styles.listType}>{title}</span>
              </div>
              <span className={styles.listDate}>{event.type === 'certification' ? 'Congratulations!' : this._formatDate(event.date)}</span>
            </div>
          );
        })}
      </div>
    );
  }

  private _renderCarouselView(): React.ReactElement {
    if (this.state.events.length === 0) {
      return <div className={styles.noEvents}>No upcoming celebrations</div>;
    }

    const event = this.state.events[this.state.currentIndex];
    return (
      <div className={styles.carouselView}>
        {this._renderEventCard(event)}
        <div className={styles.carouselIndicators}>
          {this.state.events.map((_, index) => (
            <span
              key={index}
              className={`${styles.indicator} ${index === this.state.currentIndex ? styles.active : ''}`}
              onClick={() => this.setState({ currentIndex: index })}
            />
          ))}
        </div>
      </div>
    );
  }

  public render(): React.ReactElement<ICompanyAnniversariesBirthdaysProps> {
    const { displayMode, hasTeamsContext, showTitle, centerContent } = this.props;
    const { loading, error, events } = this.state;

    return (
      <section className={`${styles.companyAnniversariesBirthdays} ${hasTeamsContext ? styles.teams : ''} ${centerContent ? styles.centered : ''}`}>
        {showTitle && (
          <div className={styles.header}>
            <h2 className={styles.title}>Company Celebrations</h2>
          </div>
        )}

        {loading && (
          <div className={styles.loading}>Loading celebrations...</div>
        )}

        {error && (
          <div className={styles.error}>{error}</div>
        )}

        {!loading && !error && events.length === 0 && (
          <div className={styles.noEvents}>No upcoming celebrations for the selected period.</div>
        )}

        {!loading && !error && events.length > 0 && (
          <>
            {displayMode === DisplayMode.Grid && this._renderGridView()}
            {displayMode === DisplayMode.List && this._renderListView()}
            {displayMode === DisplayMode.Carousel && this._renderCarouselView()}
          </>
        )}
      </section>
    );
  }
}
