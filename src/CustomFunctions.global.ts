function pad_(num) {
  return `${num < 10 ? '0' : ''}${num}`;
}

/**
 * RFC 5545 TEXT data type
 * @param {string} text Text
 * @customfunction
 */
export function ICS_TEXT(text: string) {
  text = text.toString();
  return text.replace(/([;,\\])/g, '\\$1').replace(/\n/g, '\\n');
}

type ICS_DATE_Tuple = [number, number, number];
/**
 * RFC 5545 DATE data tyoe
 * @param {string|number|ICS_DATE_Tuple} formattedTextOrYearOrArgs Date formatted per RFC 5545
 * @param {number?} month Month
 * @param {number?} day Day
 * @return {string} Formatted date
 * @customfunction
 */
export function ICS_DATE(date: string): string;
export function ICS_DATE(year: number, month: number, day: number): string;
export function ICS_DATE(args: ICS_DATE_Tuple): string;
export function ICS_DATE(
  ...args: string[] | ICS_DATE_Tuple | ICS_DATE_Tuple[]
) {
  // TODO validate params
  if (args.length && Array.isArray(args[0])) {
    if (typeof args[0] === 'string') {
      return args[0];
    } else {
      args = args[0];
    }
  }

  const [year, month, day] = args;

  return `${year}${pad_(month)}${pad_(day)}`;
}

type ICS_TIME_Tuple = [number, number, number, (boolean | string)?];
/**
 * RFC 5545 TIME data type
 * @param {string|number|ICS_TIME_Tuple} formattedTextOrHourOrArgs Time formatted per RFC 554
 * @param {number} minute Minute
 * @param {number} second Second
 * @param {boolean|string} utcOrTimeZoneId UTC (true)
 * @return {string} Formatted time
 * @customfunction
 */
export function ICS_TIME(time: string): string;
export function ICS_TIME(
  hour: number,
  minute: number,
  second: number,
  utc?: boolean
): string;
export function ICS_TIME(
  hour: number,
  minute: number,
  second: number,
  timeZoneId?: string
): string;
export function ICS_TIME(args: ICS_TIME_Tuple): string;
export function ICS_TIME(
  ...args: string[] | ICS_TIME_Tuple | ICS_TIME_Tuple[]
) {
  // TODO validate params
  if (args.length && Array.isArray(args[0])) {
    if (typeof args[0] === 'string') {
      return args[0];
    } else {
      args = args[0];
    }
  }
  const [hour, minute, second, utcOrTzid] = args;

  return `${
    utcOrTzid && utcOrTzid !== 'Z' && utcOrTzid !== true
      ? `TZID=${utcOrTzid}:`
      : ''
  }${pad_(hour)}${pad_(minute)}${pad_(second)}${
    utcOrTzid && (utcOrTzid === 'Z' || utcOrTzid === true) ? 'Z' : ''
  }`;
}

type ICS_DATE_TIME_Tuple =
  | [number, number, number]
  | [number, number, number, number, number, number, (boolean | string)?];
/**
 * date-time
 * @param {string|number|ICS_DATE_TIME_Tuple} dateTime
 * @param {number} month
 * @param {number} day
 * @param {number} hour
 * @param {number} minute
 * @param {number} second
 * @param {boolean|string} utcOrTzid
 * @returns {string}
 * @customfunction
 */
export function ICS_DATE_TIME(dateTime: string): string;
export function ICS_DATE_TIME(
  year: number,
  month: number,
  day: number,
  hour: number,
  minute: number,
  second: number,
  utc?: boolean
): string;
export function ICS_DATE_TIME(args: ICS_DATE_TIME_Tuple): string;
export function ICS_DATE_TIME(
  ...args: string[] | ICS_DATE_TIME_Tuple | ICS_DATE_TIME_Tuple[]
) {
  if (args.length && Array.isArray(args[0])) {
    if (typeof args[0] === 'string') {
      return args[0];
    } else {
      args = args[0];
    }
  }

  const [year, month, day, hour, minute, second, utcOrTzid] = ((a) =>
    a as ICS_DATE_TIME_Tuple)(args);

  return `${
    utcOrTzid && utcOrTzid !== 'Z' && utcOrTzid !== true
      ? `TZID=${utcOrTzid}:`
      : ''
  }${ICS_DATE(year, month, day)}T${ICS_TIME(
    hour,
    minute,
    second,
    utcOrTzid && (utcOrTzid === 'Z' || utcOrTzid === true) ? true : undefined
  )}`;
}

/**
 * VEVENT
 * @param {string|ICS_DATE_TIME_Tuple} dtstamp,
 * @param {string} uid,
 * @param {string|ICS_DATE_TIME_Tuple} dtstart,
 * @param {string} className
 * @param {string} created
 * @param {string} description
 * @param {string} geo
 * @param {string} lastMod
 * @param {string} location
 * @param {string} organizer
 * @param {string} priority
 * @param {string} seq
 * @param {string} status
 * @param {string} summary
 * @param {string} transp
 * @param {string} url
 * @param {string} recurid
 * @param {string} rrule
 * @param {string|ICS_DATE_TIME_Tuple} dtend
 * @param {string} duration
 * @param {string[]} attach
 * @param {string[]} vattendee
 * @param {string[]} categories
 * @param {string[]} comment
 * @param {string[]} contact
 * @param {string[]} exdate
 * @param {string[]} rstatus
 * @param {string[]} related
 * @param {string[]} resources
 * @param {string[]} rdate
 * @param {string[]} xProp
 * @customfunction
 */
export function VEVENT(
  dtstamp: string | ICS_DATE_TIME_Tuple,
  uid: string,
  dtstart: string | ICS_DATE_TIME_Tuple,
  className = undefined,
  created = undefined,
  description = undefined,
  geo = undefined,
  lastMod = undefined,
  location = undefined,
  organizer = undefined,
  priority = undefined,
  seq = undefined,
  status = undefined,
  summary = undefined,
  transp = undefined,
  url = undefined,
  recurid = undefined,
  rrule = undefined,
  dtend = undefined,
  duration = undefined,
  attach = [],
  attendee = [],
  categories = [],
  comment = [],
  contact = [],
  exdate = [],
  rstatus = [],
  related = [],
  resources = [],
  rdate = [],
  xProp = []
): string {
  if (dtend && duration) {
    throw new Error('Only one of DTEND and DURATION may be specified');
  }
  // TODO validate params
  if (Array.isArray(dtstamp)) {
    dtstamp = ICS_DATE_TIME(dtstamp as ICS_DATE_TIME_Tuple);
  }
  if (Array.isArray(dtstart)) {
    dtstart = ICS_DATE_TIME(dtstart);
  }
  const props = [];

  const addProp = (name, value, transform = (t) => t) => {
    if (value !== undefined) {
      if (value.toString().length > 0) {
        // TODO is this safe to exclude empty strings?
        value = transform(value);
        var separator = ':';
        switch (transform) {
          case ICS_DATE_TIME:
            // TODO does this logic adhere to RFC 5545?
            if (/^[^0-9]/.test(value)) {
              separator = ';';
            }
            break;
        }
        props.push(`${name}${separator}${value}`);
      }
    }
  };

  const addProps = (name, value, transform = (t) => t) => {
    if (Array.isArray(value)) {
      value.forEach((v) => addProp(name, v, transform));
    } else {
      addProp(name, value, transform);
    }
  };

  addProp('UID', uid, ICS_TEXT);
  addProp('DTSTAMP', dtstamp, ICS_DATE_TIME);
  addProp('DTSTART', dtstart, ICS_DATE_TIME);
  addProp('CLASS', className);
  addProp('CREATED', created);
  addProp('DESCRIPTION', description, ICS_TEXT);
  addProp('GEO', geo);
  addProp('LAST-MOD', lastMod, ICS_DATE_TIME);
  addProp('LOCATION', location);
  addProp('ORGANIZER', organizer);
  addProp('PRIORITY', priority);
  addProp('SEQ', seq);
  addProp('STATUS', status);
  addProp('SUMMARY', summary, ICS_TEXT);
  addProp('TRANSP', transp);
  addProp('URL', url);
  addProp('RECURID', recurid);
  addProp('RRULE', rrule);
  addProp('DTEND', dtend, ICS_DATE_TIME);
  addProp('DURATION', duration);
  addProps('ATTACH', attach);
  addProps('ATTENDEE', attendee);
  addProps('CATEGORIES', categories);
  addProps('COMMENT', comment, ICS_TEXT);
  addProps('CONTACT', contact);
  addProps('EXDATE', exdate);
  addProps('RSTATUS', rstatus);
  addProps('RELATED', related);
  addProps('RESOURCES', resources);
  addProps('RDATE', rdate);
  addProps('X-PROP', xProp);

  return ['BEGIN:VEVENT', ...props, 'END:VEVENT'].join('\n');
}

/**
 * VCALENDAR
 * @param {string} name
 * @param {string[]} body
 * @param {string} prodid
 * @customfunction
 */
export function VCALENDAR(
  name,
  body = [],
  prodid = 'Generated by Google Sheets'
) {
  return [
    'BEGIN:VCALENDAR',
    `PRODID:${ICS_TEXT(prodid)}`,
    'VERSION:2.0',
    `X-WR-CALNAME:${ICS_TEXT(name)}`,
    `X-WR-TIMEZONE:America/New_York
BEGIN:VTIMEZONE
TZID:America/New_York
X-LIC-LOCATION:America/New_York
BEGIN:DAYLIGHT
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
TZNAME:EDT
DTSTART:19700308T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:-0400
TZOFFSETTO:-0500
TZNAME:EST
DTSTART:19701101T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
END:STANDARD
END:VTIMEZONE`,
    ...body.filter((elt) => elt.toString().length > 0),
    'END:VCALENDAR'
  ];
}

/**
 * ICS feed
 * @param {string} id
 * @param {string} a1Range
 * @param {string} filename
 * @customfunction
 */
export function ICS_WEBCAL_FEED(id, a1Range, filename = undefined) {
  id = id.replace(/\./g, '_');
  const identifier = `${SpreadsheetApp.getActive().getId()}.${id}`;
  PropertiesService.getUserProperties().setProperty(
    `feed.${identifier}`,
    a1Range.toString()
  );
  filename = filename || `${id}.ics`;
  return `${process.env.SCRIPT_URL}?feed=${identifier}&filename=${filename}`;
}
