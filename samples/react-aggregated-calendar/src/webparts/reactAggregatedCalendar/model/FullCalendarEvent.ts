import * as moment from 'moment';

/**
 * Interface for FullCalendarEvent
 *
 * @export
 * @interface FullCalendarEvent
 */
export interface FullCalendarEvent {
  id: number;
  title: string;
  attorneys: any;
  start: moment.Moment;
  end: moment.Moment;
  color: string;
  allDay: boolean;
  recurrence: boolean;
  description: string;
  location: string;
  category: string;
}
