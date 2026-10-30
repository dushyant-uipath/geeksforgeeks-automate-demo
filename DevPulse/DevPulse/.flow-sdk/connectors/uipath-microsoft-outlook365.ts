import { descriptor, triggerDescriptor, type ConnectorLookupValue, type ConnectorValue } from "@uipath/maestro-builder-sdk";

export interface LocationInput {
  /** Location's unique ID, should be same as display name for online meeting. */
  uniqueId?: ConnectorValue<string>;
  /** The name displayed for the event location. For online meeting, add join URL here. */
  displayName?: ConnectorValue<string>;
}

export interface StartInput {
  /** The date and time when the event is scheduled to start */
  dateTime: ConnectorValue<string>;
  /** The time zone in which the event starts. */
  timeZone?: ConnectorValue<string>;
}

export interface BodyInput {
  /** Specifies the format of the content in the message body. */
  contentType?: ConnectorValue<string>;
  /** The description of event. */
  content?: ConnectorValue<string>;
}

export interface EndInput {
  /** The time zone in which the event ends. */
  timeZone?: ConnectorValue<string>;
  /** Specifies the date and time when the event ends */
  dateTime: ConnectorValue<string>;
}

export interface MessageInput {
  /** The primary recipients of the email, separated by comma (,) */
  toRecipients: ConnectorValue<string>;
  /** The new subject of the email. If left blank the original subject is used */
  subject?: ConnectorValue<string>;
  /** The secondary recipients of the email, separated by comma (,) */
  ccRecipients?: ConnectorValue<string>;
  /** The hidden recipients of the email, separated by comma (,) */
  bccRecipients?: ConnectorValue<string>;
}

export interface MessageInput2 {
  /** Additional primary recipients of the email, separated by comma (,) */
  toRecipients?: ConnectorValue<string>;
  /** The new subject of the email. If left blank the original subject is used */
  subject?: ConnectorValue<string>;
  /** Additional secondary recipients of the email, separated by comma (,) */
  ccRecipients?: ConnectorValue<string>;
  /** Additional hidden recipients of the email, separated by comma (,) */
  bccRecipients?: ConnectorValue<string>;
  /** The importance of the mail. */
  importance?: ConnectorValue<string>;
}

export interface MessageInput3 {
  /** The importance of the email */
  importance?: ConnectorValue<string>;
  /** The secondary recipients of the email, separated by comma (,) */
  ccRecipients?: ConnectorValue<string>;
  body?: ConnectorValue<BodyInput>;
  /** The hidden recipients of the email, separated by comma (,) */
  bccRecipients?: ConnectorValue<string>;
  /** The main recipients of the email, separated by comma(,) */
  toRecipients: ConnectorValue<string>;
  /** The subject of the email */
  subject?: ConnectorValue<string>;
  /** The email addresses to use when replying, separated by comma (,) */
  replyTo?: ConnectorValue<string>;
}

export interface BodyInput2 {
  [key: string]: unknown;
}

export interface AutomaticRepliesSettingInput {
  /** Specifies who receives automatic replies outside the organization. */
  externalAudience?: ConnectorValue<boolean>;
  scheduledEndDateTime?: ConnectorValue<ScheduledEndDateTimeInput>;
  /** The message sent automatically to external recipients. */
  externalReplyMessage?: ConnectorValue<string>;
  scheduledStartDateTime?: ConnectorValue<ScheduledStartDateTimeInput>;
  /** The message sent automatically to internal senders. */
  internalReplyMessage: ConnectorValue<string>;
  /** Indicates whether automatic replies are enabled or disabled. */
  status?: ConnectorValue<string>;
}

export interface ScheduledEndDateTimeInput {
  /** Defines the time zone for when automatic replies are scheduled to end. */
  timeZone?: ConnectorValue<string>;
  /** The date and time when automatic replies end. */
  dateTime?: ConnectorValue<string>;
}

export interface ScheduledStartDateTimeInput {
  /** Defines the time zone for when automatic replies are scheduled to start. */
  timeZone?: ConnectorValue<string>;
  /** The date and time when automatic replies begin. */
  dateTime?: ConnectorValue<string>;
}

export interface LocationInput2 {
  /** The new location of the event. If left blank, the existing value will not be updated. */
  displayName?: ConnectorValue<string>;
}

export interface BodyInput3 {
  /** The description of the event. If left blank, the existing value will not be updated. */
  content?: ConnectorValue<string>;
}

export interface StartInput2 {
  /** Start date and time for the event. */
  dateTime?: ConnectorValue<string>;
}

export interface EndInput2 {
  /** End date and time for the event. */
  dateTime?: ConnectorValue<string>;
}

export interface Attachment {
  MIMEType?: string;
  ID?: string;
  Name?: string;
  Size?: number;
}

export interface Attendee {
  Type?: string;
  Response?: string;
  Name?: string;
  Email?: string;
  status?: Status;
  /** Specifies the role of the attendee in the event. */
  type?: string;
  emailAddress?: EmailAddress;
}

export interface Category {
  Name?: string;
}

export interface Recurrence {
  range?: Range;
  pattern?: Pattern;
}

export interface Range {
  recurrenceTimeZone?: string;
  type?: string;
  numberOfOccurrences?: number;
  endDate?: string;
  startDate?: string;
}

export interface Pattern {
  type?: string;
  interval?: number;
  month?: number;
  firstDayOfWeek?: string;
  index?: string;
  dayOfMonth?: number;
  /** The days of the week when the recurrence occurs. */
  daysOfWeek?: Array<string>;
}

export interface Location {
  locationType?: string;
  displayName?: string;
  uniqueIdType?: string;
  /** A unique identifier for each location associated with the event. */
  uniqueId?: string;
}

export interface PreviousLocation {
  locationType?: string;
  displayName?: string;
  uniqueIdType?: string;
}

export interface PreviousStartDateTime {
  dateTime?: string;
  timeZone?: string;
}

export interface PreviousEndDateTime {
  dateTime?: string;
  timeZone?: string;
}

export interface From {
  emailAddress?: EmailAddress;
}

export interface EmailAddress {
  name?: string;
  address?: string;
}

export interface InternetMessageHeader {
  name?: string;
  value?: string;
}

export interface StartDateTime {
  dateTime?: string;
  timeZone?: string;
}

export interface BccRecipient {
  emailAddress?: EmailAddress;
}

export interface Flag {
  flagStatus?: string;
}

export interface Body {
  contentType?: string;
  content?: string;
}

export interface Sender {
  emailAddress?: EmailAddress;
}

export interface EndDateTime {
  dateTime?: string;
  timeZone?: string;
}

export interface CcRecipient {
  emailAddress?: EmailAddress;
}

export interface ToRecipient {
  emailAddress?: EmailAddress;
}

export interface ReplyTo {
  emailAddress?: EmailAddress;
}

export interface Status {
  /** The timestamp indicating when the attendee's status was last updated. */
  time?: string;
  /** The response status of each attendee for the event. */
  response?: string;
}

export interface ResponseStatus {
  /** The current response status of the attendee, such as accepted or declined. */
  response?: string;
  /** The time when the response status was recorded. */
  time?: string;
}

export interface OnlineMeeting {
  /** Provides the URL to join the online meeting associated with the event. */
  joinUrl?: string;
}

export interface Organizer {
  emailAddress?: EmailAddress;
}

export interface Owner {
  address?: string;
  name?: string;
}

export interface Odata {
  /** Specifies the type of data structure used in the response. */
  type?: string;
  /** A unique identifier for the version of the email entity. */
  etag?: string;
}

export interface Start {
  /** The date and time when the event is scheduled to start */
  dateTime?: string;
  /** The time zone in which the event starts. */
  timeZone?: string;
}

export interface End {
  /** The time zone in which the event ends. */
  timeZone?: string;
  /** Specifies the date and time when the event ends */
  dateTime?: string;
}

export interface AutomaticRepliesSetting {
  /** Specifies who receives automatic replies outside your organization. */
  externalAudience?: boolean | string;
  scheduledEndDateTime?: ScheduledEndDateTime;
  /** The message sent automatically to external contacts when you are unavailable. */
  externalReplyMessage?: string;
  scheduledStartDateTime?: ScheduledStartDateTime;
  /** The message sent automatically to internal contacts when you are unavailable. */
  internalReplyMessage?: string;
  /** Indicates whether automatic replies are enabled or disabled. */
  status?: string;
}

export interface ScheduledEndDateTime {
  /** Defines the time zone for when automatic replies are scheduled to end. */
  timeZone?: string;
  /** The date and time when automatic replies will stop. */
  dateTime?: string;
}

export interface ScheduledStartDateTime {
  /** Defines the time zone for when automatic replies are scheduled to start. */
  timeZone?: string;
  /** The date and time when automatic replies will start. */
  dateTime?: string;
}

export interface CalendarEventCreatedWhere {
  /** Calendar from which to retrieve events LOOKUP — sends a `CalendarList.ID`, never a name or e-mail: lookup(<Descriptor>, 'calendarID') (helpers in the LOOKUP lines above the descriptor). */
  calendarID?: ConnectorLookupValue<string>;
}

export interface CalendarEventCreatedOutputs {
  Attachments?: Array<Attachment>;
  Description?: string;
  CalendarID?: string;
  HasAttachments?: boolean;
  Title?: string;
  CalendarName?: string;
  Attendees?: Array<Attendee>;
  Importance?: string;
  Sensitivity?: string;
  AllDay?: boolean;
  Categories?: Array<Category>;
  ID?: string;
  Location?: string;
  ShowAs?: string;
}

/**
 * Calendar Event Created — uipath.connector.event.uipath-microsoft-outlook365.calendar-event-created
 *
 * Start a flow on it with `onEvent(CalendarEventCreated, { where: { … } })`, or pause for it
 * with `waitForEvent(CalendarEventCreated, { … })`; `where` carries its event parameters.
 *
 * LOOKUP: `calendarID` -> CalendarList.ID
 *   by: `FullName`
 *   use: lookup(CalendarEventCreated, 'calendarID').byFullName(value)
 *   also: .byFullName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const CalendarEventCreated = triggerDescriptor<CalendarEventCreatedWhere, CalendarEventCreatedOutputs>()({
  connector: "uipath-microsoft-outlook365",
  event: "calendar-event-created",
  version: "1.0.0",
  displayName: "Calendar Event Created",
  lookups: {
    calendarID: {
      objectName: "CalendarList",
      path: "/CalendarList",
      by: ["FullName"],
      value: "ID",
      aliases: { byFullName: "FullName" },
      strategy: "tree",
    },
  },
} as const);

export interface CalendarEventReceivedWhere {
  /** Calendar from which to retrieve events LOOKUP — sends a `CalendarList.ID`, never a name or e-mail: lookup(<Descriptor>, 'calendarID') (helpers in the LOOKUP lines above the descriptor). */
  calendarID?: ConnectorLookupValue<string>;
}

export interface CalendarEventReceivedOutputs {
  Attachments?: Array<Attachment>;
  Description?: string;
  CalendarID?: string;
  HasAttachments?: boolean;
  Title?: string;
  CalendarName?: string;
  Attendees?: Array<Attendee>;
  Importance?: string;
  Sensitivity?: string;
  AllDay?: boolean;
  Categories?: Array<Category>;
  ID?: string;
  Location?: string;
  ShowAs?: string;
}

/**
 * Calendar Event Received — uipath.connector.event.uipath-microsoft-outlook365.calendar-event-received
 *
 * Start a flow on it with `onEvent(CalendarEventReceived, { where: { … } })`, or pause for it
 * with `waitForEvent(CalendarEventReceived, { … })`; `where` carries its event parameters.
 *
 * LOOKUP: `calendarID` -> CalendarList.ID
 *   by: `FullName`
 *   use: lookup(CalendarEventReceived, 'calendarID').byFullName(value)
 *   also: .byFullName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const CalendarEventReceived = triggerDescriptor<CalendarEventReceivedWhere, CalendarEventReceivedOutputs>()({
  connector: "uipath-microsoft-outlook365",
  event: "calendar-event-received",
  version: "1.0.0",
  displayName: "Calendar Event Received",
  lookups: {
    calendarID: {
      objectName: "CalendarList",
      path: "/CalendarList",
      by: ["FullName"],
      value: "ID",
      aliases: { byFullName: "FullName" },
      strategy: "tree",
    },
  },
} as const);

export interface CalendarEventRepliedWhere {
  /** Calendar from which to retrieve events LOOKUP — sends a `CalendarList.ID`, never a name or e-mail: lookup(<Descriptor>, 'calendarID') (helpers in the LOOKUP lines above the descriptor). */
  calendarID?: ConnectorLookupValue<string>;
}

export interface CalendarEventRepliedOutputs {
  Attachments?: Array<Attachment>;
  Description?: string;
  CalendarID?: string;
  HasAttachments?: boolean;
  Title?: string;
  CalendarName?: string;
  Attendees?: Array<Attendee>;
  Importance?: string;
  Sensitivity?: string;
  AllDay?: boolean;
  Categories?: Array<Category>;
  ID?: string;
  Location?: string;
  ShowAs?: string;
}

/**
 * Calendar Event Replied — uipath.connector.event.uipath-microsoft-outlook365.calendar-event-replied
 *
 * Start a flow on it with `onEvent(CalendarEventReplied, { where: { … } })`, or pause for it
 * with `waitForEvent(CalendarEventReplied, { … })`; `where` carries its event parameters.
 *
 * LOOKUP: `calendarID` -> CalendarList.ID
 *   by: `FullName`
 *   use: lookup(CalendarEventReplied, 'calendarID').byFullName(value)
 *   also: .byFullName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const CalendarEventReplied = triggerDescriptor<CalendarEventRepliedWhere, CalendarEventRepliedOutputs>()({
  connector: "uipath-microsoft-outlook365",
  event: "calendar-event-replied",
  version: "1.0.0",
  displayName: "Calendar Event Replied",
  lookups: {
    calendarID: {
      objectName: "CalendarList",
      path: "/CalendarList",
      by: ["FullName"],
      value: "ID",
      aliases: { byFullName: "FullName" },
      strategy: "tree",
    },
  },
} as const);

export interface CalendarEventUpdatedWhere {
  /** Calendar from which to retrieve events LOOKUP — sends a `CalendarList.ID`, never a name or e-mail: lookup(<Descriptor>, 'calendarID') (helpers in the LOOKUP lines above the descriptor). */
  calendarID?: ConnectorLookupValue<string>;
}

export interface CalendarEventUpdatedOutputs {
  Attachments?: Array<Attachment>;
  Description?: string;
  CalendarID?: string;
  HasAttachments?: boolean;
  Title?: string;
  CalendarName?: string;
  Attendees?: Array<Attendee>;
  Importance?: string;
  Sensitivity?: string;
  AllDay?: boolean;
  Categories?: Array<Category>;
  ID?: string;
  Location?: string;
  ShowAs?: string;
}

/**
 * Calendar Event Updated — uipath.connector.event.uipath-microsoft-outlook365.calendar-event-updated
 *
 * Start a flow on it with `onEvent(CalendarEventUpdated, { where: { … } })`, or pause for it
 * with `waitForEvent(CalendarEventUpdated, { … })`; `where` carries its event parameters.
 *
 * LOOKUP: `calendarID` -> CalendarList.ID
 *   by: `FullName`
 *   use: lookup(CalendarEventUpdated, 'calendarID').byFullName(value)
 *   also: .byFullName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const CalendarEventUpdated = triggerDescriptor<CalendarEventUpdatedWhere, CalendarEventUpdatedOutputs>()({
  connector: "uipath-microsoft-outlook365",
  event: "calendar-event-updated",
  version: "1.0.0",
  displayName: "Calendar Event Updated",
  lookups: {
    calendarID: {
      objectName: "CalendarList",
      path: "/CalendarList",
      by: ["FullName"],
      value: "ID",
      aliases: { byFullName: "FullName" },
      strategy: "tree",
    },
  },
} as const);

export interface EmailMovedToTrashOutputs {
  lastModifiedDateTime?: string;
  recurrence?: Recurrence;
  location?: Location | Array<Location>;
  parentFolderId?: string;
  odataEtag?: string;
  subject?: string;
  webLink?: string;
  isDraft?: boolean;
  previousLocation?: PreviousLocation;
  bodyPreview?: string;
  createdDateTime?: string;
  previousStartDateTime?: PreviousStartDateTime;
  type?: string;
  previousEndDateTime?: PreviousEndDateTime;
  from?: From;
  meetingMessageType?: string;
  internetMessageHeaders?: Array<InternetMessageHeader>;
  startDateTime?: StartDateTime;
  bccRecipients?: Array<BccRecipient>;
  isReadReceiptRequested?: boolean;
  id?: string;
  sentDateTime?: string;
  odataContext?: string;
  parentFolderName?: string;
  changeKey?: string;
  receivedDateTime?: string;
  isDeliveryReceiptRequested?: boolean;
  isOutOfDate?: boolean;
  flag?: Flag;
  body?: Body;
  importance?: string;
  isRead?: boolean;
  sender?: Sender;
  categories?: Array<string>;
  endDateTime?: EndDateTime;
  ccRecipients?: Array<CcRecipient>;
  inferenceClassification?: string;
  internetMessageId?: string;
  isAllDay?: boolean;
  toRecipients?: Array<ToRecipient>;
  isDelegated?: boolean;
  replyTo?: Array<ReplyTo>;
  hasAttachments?: boolean;
  responseRequested?: boolean;
  conversationIndex?: string;
  meetingRequestType?: string;
  conversationId?: string;
}

/**
 * Email Moved to Trash — uipath.connector.event.uipath-microsoft-outlook365.email-moved-to-trash
 *
 * Start a flow on it with `onEvent(EmailMovedToTrash, { where: { … } })`, or pause for it
 * with `waitForEvent(EmailMovedToTrash, { … })`; `where` carries its event parameters.
 */
export const EmailMovedToTrash = triggerDescriptor<Record<string, string>, EmailMovedToTrashOutputs>()({
  connector: "uipath-microsoft-outlook365",
  event: "email-moved-to-trash",
  version: "1.0.0",
  displayName: "Email Moved to Trash",
} as const);

export interface EmailReceivedWhere {
  /** Email folder LOOKUP — sends a `MailFolder.id`, never a name or e-mail: lookup(<Descriptor>, 'parentFolderId') (helpers in the LOOKUP lines above the descriptor). */
  parentFolderId: ConnectorLookupValue<string>;
}

export interface EmailReceivedOutputs {
  lastModifiedDateTime?: string;
  recurrence?: Recurrence;
  location?: Location;
  parentFolderId?: string;
  odataEtag?: string;
  subject?: string;
  webLink?: string;
  isDraft?: boolean;
  previousLocation?: PreviousLocation;
  bodyPreview?: string;
  createdDateTime?: string;
  previousStartDateTime?: PreviousStartDateTime;
  type?: string;
  previousEndDateTime?: PreviousEndDateTime;
  from?: From;
  meetingMessageType?: string;
  internetMessageHeaders?: Array<InternetMessageHeader>;
  startDateTime?: StartDateTime;
  bccRecipients?: Array<BccRecipient>;
  isReadReceiptRequested?: boolean;
  id?: string;
  sentDateTime?: string;
  odataContext?: string;
  parentFolderName?: string;
  changeKey?: string;
  receivedDateTime?: string;
  isDeliveryReceiptRequested?: boolean;
  isOutOfDate?: boolean;
  flag?: Flag;
  body?: Body;
  importance?: string;
  isRead?: boolean;
  sender?: Sender;
  categories?: Array<string>;
  endDateTime?: EndDateTime;
  ccRecipients?: Array<CcRecipient>;
  inferenceClassification?: string;
  internetMessageId?: string;
  isAllDay?: boolean;
  toRecipients?: Array<ToRecipient>;
  isDelegated?: boolean;
  replyTo?: Array<ReplyTo>;
  hasAttachments?: boolean;
  responseRequested?: boolean;
  conversationIndex?: string;
  meetingRequestType?: string;
  conversationId?: string;
}

/**
 * Email Received — uipath.connector.event.uipath-microsoft-outlook365.email-received
 *
 * Start a flow on it with `onEvent(EmailReceived, { where: { … } })`, or pause for it
 * with `waitForEvent(EmailReceived, { … })`; `where` carries its event parameters.
 *
 * LOOKUP: `parentFolderId` -> MailFolder.id
 *   by: `displayName`
 *   use: lookup(EmailReceived, 'parentFolderId').byDisplayName(value)
 *   also: .byDisplayName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const EmailReceived = triggerDescriptor<EmailReceivedWhere, EmailReceivedOutputs>()({
  connector: "uipath-microsoft-outlook365",
  event: "email-received",
  version: "1.0.0",
  displayName: "Email Received",
  lookups: {
    parentFolderId: {
      objectName: "MailFolder",
      path: "/MailFolders",
      by: ["displayName"],
      value: "id",
      aliases: { byDisplayName: "displayName" },
      strategy: "tree",
    },
  },
} as const);

export interface EmailSentOutputs {
  lastModifiedDateTime?: string;
  recurrence?: Recurrence;
  location?: Location;
  parentFolderId?: string;
  odataEtag?: string;
  subject?: string;
  webLink?: string;
  isDraft?: boolean;
  previousLocation?: PreviousLocation;
  bodyPreview?: string;
  createdDateTime?: string;
  previousStartDateTime?: PreviousStartDateTime;
  type?: string;
  previousEndDateTime?: PreviousEndDateTime;
  from?: From;
  meetingMessageType?: string;
  internetMessageHeaders?: Array<InternetMessageHeader>;
  startDateTime?: StartDateTime;
  bccRecipients?: Array<BccRecipient>;
  isReadReceiptRequested?: boolean;
  id?: string;
  sentDateTime?: string;
  odataContext?: string;
  parentFolderName?: string;
  changeKey?: string;
  receivedDateTime?: string;
  isDeliveryReceiptRequested?: boolean;
  isOutOfDate?: boolean;
  flag?: Flag;
  body?: Body;
  importance?: string;
  isRead?: boolean;
  sender?: Sender;
  categories?: Array<string>;
  endDateTime?: EndDateTime;
  ccRecipients?: Array<CcRecipient>;
  inferenceClassification?: string;
  internetMessageId?: string;
  isAllDay?: boolean;
  toRecipients?: Array<ToRecipient>;
  isDelegated?: boolean;
  replyTo?: Array<ReplyTo>;
  hasAttachments?: boolean;
  responseRequested?: boolean;
  conversationIndex?: string;
  meetingRequestType?: string;
  conversationId?: string;
}

/**
 * Email Sent — uipath.connector.event.uipath-microsoft-outlook365.email-sent
 *
 * Start a flow on it with `onEvent(EmailSent, { where: { … } })`, or pause for it
 * with `waitForEvent(EmailSent, { … })`; `where` carries its event parameters.
 */
export const EmailSent = triggerDescriptor<Record<string, string>, EmailSentOutputs>()({
  connector: "uipath-microsoft-outlook365",
  event: "email-sent",
  version: "1.0.0",
  displayName: "Email Sent",
} as const);

export interface ExistingRecordUpdated__Outputs {
}

/**
 * Existing Record Updated — uipath.connector.event.uipath-microsoft-outlook365.existing-record-updated
 *
 * A GENERIC event: it fires for ONE object of the connection (an entity, a
 * table, a custom object), so name it — `onEvent(ExistingRecordUpdated__, { object: '<Object>' })`.
 * `uip is triggers objects uipath-microsoft-outlook365 <EVENT> --connection-id <id>` lists them.
 */
export const ExistingRecordUpdated__ = triggerDescriptor<Record<string, string>, ExistingRecordUpdated__Outputs>()({
  connector: "uipath-microsoft-outlook365",
  event: "existing-record-updated",
  version: "1.0.0",
  displayName: "Existing Record Updated",
} as const);

export interface NewRecordCreated__Outputs {
}

/**
 * New Record Created — uipath.connector.event.uipath-microsoft-outlook365.new-record-created
 *
 * A GENERIC event: it fires for ONE object of the connection (an entity, a
 * table, a custom object), so name it — `onEvent(NewRecordCreated__, { object: '<Object>' })`.
 * `uip is triggers objects uipath-microsoft-outlook365 <EVENT> --connection-id <id>` lists them.
 */
export const NewRecordCreated__ = triggerDescriptor<Record<string, string>, NewRecordCreated__Outputs>()({
  connector: "uipath-microsoft-outlook365",
  event: "new-record-created",
  version: "1.0.0",
  displayName: "New Record Created",
} as const);

export interface EmailReceived_Outputs {
  /** The email's subject line. */
  subject?: string;
  /** The id of the mail folder the email arrived in. */
  parentFolderId?: string;
  /** The display name of the mail folder the email arrived in. */
  parentFolderName?: string;
  from?: From;
  /** The first part of the email's body, as plain text. */
  bodyPreview?: string;
  /** The importance the sender set on the email (low | normal | high). */
  importance?: string;
}

/**
 * Email Received — uipath.connector.trigger.uipath-microsoft-outlook365.email-received
 *
 * Start a flow on it with `onEvent(EmailReceived_, { where: { … } })`, or pause for it
 * with `waitForEvent(EmailReceived_, { … })`; `where` carries its event parameters.
 */
export const EmailReceived_ = triggerDescriptor<Record<string, string>, EmailReceived_Outputs>()({
  connector: "uipath-microsoft-outlook365",
  event: "email-received",
  version: "1.0.0",
  displayName: "Email Received",
} as const);

export interface ArchiveEmailInputs {
  /** The ID of email message to archive */
  id: ConnectorValue<string>;
}

export interface ArchiveEmailOutputs {
  /** The date and time when the email was last modified. */
  lastModifiedDateTime?: string;
  body?: Body;
  /** The unique identifier of the folder containing the email. */
  parentFolderId?: string;
  /** The subject line of the email. */
  subject?: string;
  /** The priority level of the email, such as high, normal, or low. */
  importance?: string;
  /** A URL link to access the email in a web browser. */
  webLink?: string;
  /** Indicates whether the email has been read. */
  isRead?: boolean;
  /** Indicates whether the email is a draft. */
  isDraft?: boolean;
  /** The date and time when the email was created. */
  createdDateTime?: string;
  /** A short preview of the email's content. */
  bodyPreview?: string;
  sender?: Sender;
  /** The classification of the email based on inferred importance. */
  inferenceClassification?: string;
  from?: From;
  /** A unique identifier for the email message on the internet. */
  internetMessageId?: string;
  toRecipients?: Array<ToRecipient>;
  /** Indicates if a read receipt is requested. */
  isReadReceiptRequested?: boolean;
  /** The unique identifier for the email. */
  id?: string;
  /** Indicates if the email has attachments. */
  hasAttachments?: boolean;
  /** The date and time when the email was sent. */
  sentDateTime?: string;
  /** A unique identifier for the email's position in a thread. */
  conversationIndex?: string;
  /** A unique identifier for the email conversation thread. */
  conversationId?: string;
  /** Identifies the version of the email item. */
  changeKey?: string;
  /** Indicates when the email was received. */
  receivedDateTime?: string;
  flag?: Flag;
}

/** Archive Email — uipath.connector.uipath-microsoft-outlook365.archive-email */
export const ArchiveEmail = descriptor<ArchiveEmailInputs, ArchiveEmailOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.archive-email",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface CreateEventCalendarEventInputs {
  /** Title → e.g. Event title */
  subject: ConnectorValue<string>;
  /** Indicates if attendees can propose new times for the event. */
  allowNewTimeProposals?: ConnectorValue<boolean>;
  /** Comma separated list of optional attendees. */
  optionalAttendees?: ConnectorValue<string>;
  location?: ConnectorValue<LocationInput>;
  /** Indicates how the event appears on the calendar, like busy or free. */
  showAs?: ConnectorValue<string>;
  /** Comma separated list of resources, like rooms or equipment, invited to the event. */
  resourceAttendees?: ConnectorValue<string>;
  start: ConnectorValue<StartInput>;
  /** Indicates the privacy level of the calendar event. */
  sensitivity?: ConnectorValue<string>;
  /** Indicates if the meeting is set as an online meeting. Should only be marked as true for Teams meeting. */
  isOnlineMeeting?: ConnectorValue<boolean>;
  body?: ConnectorValue<BodyInput>;
  /** Defines the priority level of the calendar event. */
  importance?: ConnectorValue<string>;
  end: ConnectorValue<EndInput>;
  /** The categories to set on the event */
  categories?: ConnectorValue<Array<string>>;
  /** Indicates if the event lasts the entire day. */
  isAllDay?: ConnectorValue<boolean>;
  /** Comma separated list of required attendees. */
  requiredAttendees?: ConnectorValue<string>;
  /** Indicates if attendees are hidden from the calendar event. */
  hideAttendees?: ConnectorValue<boolean>;
  /** Chose or type a value */
  timezone: ConnectorValue<string>;
  /** Attachment for event */
  file?: ConnectorValue<string>;
  /** Chose a value */
  outputTimezone?: ConnectorValue<string>;
  /** Calendar to create event LOOKUP — sends a `CalendarList.ID`, never a name or e-mail: lookup(<Descriptor>, 'calendarID') (helpers in the LOOKUP lines above the descriptor). */
  calendarID?: ConnectorLookupValue<string>;
}

export interface CreateEventCalendarEventOutputs {
  /** The date and time when the calendar entry was last updated. */
  lastModifiedDateTime?: string;
  location?: Location;
  locations?: Array<Location>;
  /** A URL link to access the calendar event online. */
  webLink?: string;
  /** Indicates if the calendar event is saved as a draft. */
  isDraft?: boolean;
  /** A short preview of the event's description or content. */
  bodyPreview?: string;
  /** The date and time when the calendar event was created. */
  createdDateTime?: string;
  /** Specifies the type of calendar event or action. */
  type?: string;
  /** The time zone in which the event originally ends. */
  originalEndTimeZone?: string;
  /** The number of minutes before the event to send a reminder. */
  reminderMinutesBeforeStart?: number;
  attendees?: Array<Attendee>;
  /** A unique identifier for the calendar event. */
  id?: string;
  /** The time zone in which the event was originally scheduled. */
  originalStartTimeZone?: string;
  responseStatus?: ResponseStatus;
  /** A unique identifier for tracking the calendar action transaction. */
  transactionId?: string;
  /** A unique identifier for tracking changes to the event. */
  changeKey?: string;
  /** Indicates whether the user is the organizer of the event. */
  isOrganizer?: boolean;
  onlineMeeting?: OnlineMeeting;
  /** A unique identifier for the calendar event in iCalendar format. */
  iCalUId?: string;
  organizer?: Organizer;
  /** The service used for hosting the online meeting. */
  onlineMeetingProvider?: string;
  /** Indicates whether the calendar event includes attachments. */
  hasAttachments?: boolean;
  /** Indicates if a response is requested from attendees. */
  responseRequested?: boolean;
  /** Shows whether the event has been cancelled. */
  isCancelled?: boolean;
  /** Indicates if a reminder is set for the event. */
  isReminderOn?: boolean;
  /** The web link to join the online meeting associated with the event. */
  onlineMeetingUrl?: string;
}

/**
 * Create Event (Calendar Event) — uipath.connector.uipath-microsoft-outlook365.create-event-calendar-event-
 *
 * LOOKUP: `calendarID` -> CalendarList.ID
 *   by: `FullName`
 *   use: lookup(CreateEventCalendarEvent, 'calendarID').byFullName(value)
 *   also: .byFullName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const CreateEventCalendarEvent = descriptor<CreateEventCalendarEventInputs, CreateEventCalendarEventOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.create-event-calendar-event-",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
  lookups: {
    calendarID: {
      objectName: "CalendarList",
      path: "/CalendarList",
      by: ["FullName"],
      value: "ID",
      aliases: { byFullName: "FullName" },
      strategy: "tree",
    },
  },
} as const);

export interface DeleteEmailInputs {
  /** The ID of email to delete */
  id: ConnectorValue<string>;
  /** Indicates whether to delete the email permanently */
  permanentlyDelete?: ConnectorValue<boolean>;
}

export interface DeleteEmailOutputs {
}

/** Delete Email — uipath.connector.uipath-microsoft-outlook365.delete-email */
export const DeleteEmail = descriptor<DeleteEmailInputs, DeleteEmailOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.delete-email",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface DeleteEventCalendarEventInputs {
  /** The ID of event to cancel/delete */
  id: ConnectorValue<string>;
  /** The message sent to the attendees/organizer of the deleted event */
  comment?: ConnectorValue<string>;
  /** The ID of calendar in which respective event is present. Should be left empty for personal calendar. LOOKUP — sends a `CalendarList.ID`, never a name or e-mail: lookup(<Descriptor>, 'calendarID') (helpers in the LOOKUP lines above the descriptor). */
  calendarID?: ConnectorLookupValue<string>;
  /** Indicates option for event deletion */
  deleteOption?: ConnectorValue<string>;
}

export interface DeleteEventCalendarEventOutputs {
}

/**
 * Delete Event (Calendar Event) — uipath.connector.uipath-microsoft-outlook365.delete-event-calendar-event-
 *
 * LOOKUP: `calendarID` -> CalendarList.ID
 *   by: `FullName`
 *   use: lookup(DeleteEventCalendarEvent, 'calendarID').byFullName(value)
 *   also: .byFullName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const DeleteEventCalendarEvent = descriptor<DeleteEventCalendarEventInputs, DeleteEventCalendarEventOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.delete-event-calendar-event-",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
  lookups: {
    calendarID: {
      objectName: "CalendarList",
      path: "/CalendarList",
      by: ["FullName"],
      value: "ID",
      aliases: { byFullName: "FullName" },
      strategy: "tree",
    },
  },
} as const);

export interface DownloadEmailInputs {
  /** The ID of email you want to download */
  id: ConnectorValue<string>;
}

export interface DownloadEmailOutputs {
  /** Final output file */
  Response?: unknown;
}

/** Download Email — uipath.connector.uipath-microsoft-outlook365.download-email */
export const DownloadEmail = descriptor<DownloadEmailInputs, DownloadEmailOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.download-email",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface DownloadEmailAttachmentInputs {
  /** Download only the attachment whose name matches the specified pattern. Multiple filename patterns should be separated by a vertical bar ('|') character. A pattern may contain the wildcard character ('*'), which matches zero or more characters (eg. *.xlsx or file1.pdf or fi*.pdf|*.docx) */
  fileName?: ConnectorValue<string>;
  /** The ID of the email from which you want to download the attachment */
  id: ConnectorValue<string>;
  /** The ID of the attachment to download, as returned by List Email Attachments. Use this to pick an exact attachment when several share the same name. When provided, the filename filter is ignored; the exclude inline attachments option still applies */
  attachmentId?: ConnectorValue<string>;
  /** Indicates whether to exclude the inline attachments, i.e. attachments embedded in the email body */
  excludeInlineAttachments?: ConnectorValue<boolean>;
}

export interface DownloadEmailAttachmentOutputs {
  /** Final output file */
  Response?: unknown;
}

/** Download Email Attachment — uipath.connector.uipath-microsoft-outlook365.download-email-attachment */
export const DownloadEmailAttachment = descriptor<DownloadEmailAttachmentInputs, DownloadEmailAttachmentOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.download-email-attachment",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface ForwardEmailInputs {
  message: ConnectorValue<MessageInput>;
  /** The body of the email */
  comment?: ConnectorValue<string>;
  /** The attachment to be sent with the email */
  file?: ConnectorValue<string>;
  /** The ID of email to forward */
  id: ConnectorValue<string>;
  /** Specifies whether the email should be saved as draft instead of being sent */
  saveAsDraft?: ConnectorValue<boolean>;
  /** Set the timezone for forwarded email messages. If not specified, the default timezone will be UTC. */
  timezone?: ConnectorValue<string>;
}

export interface ForwardEmailOutputs {
  /** The status of the forward operation. */
  status?: string;
  /** A unique identifier for the forwarded message. */
  id?: string;
  /** Unique identifier for the email conversation thread. Use this with the List Emails / Get Newest Email Conversation ID filter to track incoming replies. */
  conversationId?: string;
  /** A unique identifier for the email message on the internet. */
  internetMessageId?: string;
}

/** Forward Email — uipath.connector.uipath-microsoft-outlook365.forward-email */
export const ForwardEmail = descriptor<ForwardEmailInputs, ForwardEmailOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.forward-email",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface ForwardEventCalendarEventInputs {
  /** A comma separated email of additional attendees receiving the event. */
  to: ConnectorValue<string>;
  /** The message sent to the recipients of the forwarded event along the invitation */
  comment?: ConnectorValue<string>;
  /** The event to forward */
  id: ConnectorValue<string>;
  /** The calendar in which respective event is present LOOKUP — sends a `CalendarList.ID`, never a name or e-mail: lookup(<Descriptor>, 'calendarID') (helpers in the LOOKUP lines above the descriptor). */
  calendarID?: ConnectorLookupValue<string>;
  /** Forward all occurances of given event */
  applyOnSeries?: ConnectorValue<boolean>;
}

export interface ForwardEventCalendarEventOutputs {
}

/**
 * Forward Event (Calendar Event) — uipath.connector.uipath-microsoft-outlook365.forward-event-calendar-event-
 *
 * LOOKUP: `calendarID` -> CalendarList.ID
 *   by: `FullName`
 *   use: lookup(ForwardEventCalendarEvent, 'calendarID').byFullName(value)
 *   also: .byFullName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const ForwardEventCalendarEvent = descriptor<ForwardEventCalendarEventInputs, ForwardEventCalendarEventOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.forward-event-calendar-event-",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
  lookups: {
    calendarID: {
      objectName: "CalendarList",
      path: "/CalendarList",
      by: ["FullName"],
      value: "ID",
      aliases: { byFullName: "FullName" },
      strategy: "tree",
    },
  },
} as const);

export interface GetCalendarsInputs {
  /** The CEQL search expression */
  where?: ConnectorValue<string>;
  /** The number of resources to return in a given page */
  pageSize?: ConnectorValue<number>;
  /** The next page cursor, taken from the response header: `elements-next-page-token` */
  nextPage?: ConnectorValue<string>;
}

export interface GetCalendarsOutputs {
  allowedOnlineMeetingProviders?: Array<string>;
  canEdit?: boolean;
  canShare?: boolean;
  canViewPrivateItems?: boolean;
  changeKey?: string;
  color?: string;
  defaultOnlineMeetingProvider?: string;
  hexColor?: string;
  id?: string;
  isDefaultCalendar?: boolean;
  isRemovable?: boolean;
  isTallyingResponses?: boolean;
  name?: string;
  odataId?: string;
  owner?: Owner;
}

/** Get Calendars — uipath.connector.uipath-microsoft-outlook365.get-calendars */
export const GetCalendars = descriptor<GetCalendarsInputs, GetCalendarsOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.get-calendars",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface GetEmailByIdInputs {
  /** List of fields to select */
  fields?: ConnectorValue<string>;
  /** The unique identifier for the email */
  id: ConnectorValue<string>;
}

export interface GetEmailByIdOutputs {
  bccRecipients?: Array<BccRecipient>;
  body?: Body;
  bodyPreview?: string;
  categories?: Array<string>;
  ccRecipients?: Array<CcRecipient>;
  changeKey?: string;
  conversationId?: string;
  conversationIndex?: string;
  createdDateTime?: string;
  endDateTime?: EndDateTime;
  flag?: Flag;
  from?: From;
  hasAttachments?: boolean;
  id?: string;
  importance?: string;
  inferenceClassification?: string;
  internetMessageId?: string;
  isAllDay?: boolean;
  isDelegated?: boolean;
  isDeliveryReceiptRequested?: boolean;
  isDraft?: boolean;
  isOutOfDate?: boolean;
  isRead?: boolean;
  isReadReceiptRequested?: boolean;
  lastModifiedDateTime?: string;
  location?: Location;
  meetingMessageType?: string;
  meetingRequestType?: string;
  parentFolderId?: string;
  parentFolderName?: string;
  previousEndDateTime?: PreviousEndDateTime;
  previousLocation?: PreviousLocation;
  previousStartDateTime?: PreviousStartDateTime;
  receivedDateTime?: string;
  recurrence?: Recurrence;
  replyTo?: Array<ReplyTo>;
  responseRequested?: boolean;
  sender?: Sender;
  sentDateTime?: string;
  startDateTime?: StartDateTime;
  subject?: string;
  toRecipients?: Array<ToRecipient>;
  type?: string;
  webLink?: string;
}

/** Get Email by ID — uipath.connector.uipath-microsoft-outlook365.get-email-by-id */
export const GetEmailById = descriptor<GetEmailByIdInputs, GetEmailByIdOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.get-email-by-id",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface GetEmailFoldersListInputs {
  /** The number of resources to return in a given page */
  pageSize?: ConnectorValue<number>;
  /** The next page cursor, taken from the response header: `elements-next-page-token` */
  nextPage?: ConnectorValue<string>;
  /** The next page token, taken from the response headers 'elements-next-page-token' */
  page?: ConnectorValue<string>;
  /** The OData query parameter to retrieve a subset of the MailFolders */
  filter?: ConnectorValue<string>;
  /** The OData query parameter to order the MailFolders by */
  orderby?: ConnectorValue<string>;
  /** The Parent folder ID to get the child folders */
  parentFolderId?: ConnectorValue<string>;
  /** Shared mailbox's email address for polling */
  sharedMailboxAddress?: ConnectorValue<string>;
  /** Set to true to recursively include all nested subfolders in the response. Defaults to false (top-level folders only). */
  includeChildFolders?: ConnectorValue<boolean>;
}

export interface GetEmailFoldersListOutputs {
  childFolderCount?: number;
  displayName?: string;
  id?: string;
  isHidden?: boolean;
  parentFolderId?: string;
  totalItemCount?: number;
  unreadItemCount?: number;
  isFolder?: boolean;
}

/** Get Email Folders List — uipath.connector.uipath-microsoft-outlook365.get-email-folders-list */
export const GetEmailFoldersList = descriptor<GetEmailFoldersListInputs, GetEmailFoldersListOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.get-email-folders-list",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface GetEmailListInputs {
  /** Indicates whether to consider only unread emails or not */
  unReadOnly?: ConnectorValue<boolean>;
  /** Additional OData filters */
  filter?: ConnectorValue<string>;
  /** The number of emails to retrieve */
  limit?: ConnectorValue<string>;
  /** Importance of the email */
  importance?: ConnectorValue<string>;
  /** The folder to get the emails from LOOKUP — sends a `MailFolder.id`, never a name or e-mail: lookup(<Descriptor>, 'parentFolderId') (helpers in the LOOKUP lines above the descriptor). */
  parentFolderId: ConnectorLookupValue<string>;
  /** Indicates whether to consider only emails with attachments or not */
  withAttachmentsOnly?: ConnectorValue<boolean>;
  /** Indicates whether to expand the search to include all subfolders of the selected mail folder or not */
  includeSubfolders?: ConnectorValue<boolean>;
  /** Indicates whether to mark the retrieved email as read */
  markAsRead?: ConnectorValue<boolean>;
  /** Return only emails in this conversation thread. Use the Conversation Id returned by Reply To Email (or a message trigger) to track incoming replies. */
  conversationId?: ConnectorValue<string>;
}

export interface GetEmailListOutputs {
  /** The date and time when the email was last modified. */
  lastModifiedDateTime?: string;
  recurrence?: Recurrence;
  location?: Location;
  /** The unique identifier of the parent folder containing the email. */
  parentFolderId?: string;
  /** The subject line of the email message. */
  subject?: string;
  /** A URL that directs you to view the email online. */
  webLink?: string;
  odata?: Odata;
  /** Indicates whether the email is saved as a draft. */
  isDraft?: boolean;
  /** A short preview of the email's content. */
  bodyPreview?: string;
  /** The date and time when the email was created. */
  createdDateTime?: string;
  previousStartDateTime?: PreviousStartDateTime;
  /** Specifies the type of event or action being performed. */
  type?: string;
  previousEndDateTime?: PreviousEndDateTime;
  /** Describes the type of message related to the meeting. */
  meetingMessageType?: string;
  from?: From;
  startDateTime?: StartDateTime;
  /** Indicates if a read receipt is requested for the email. */
  isReadReceiptRequested?: boolean;
  /** A unique identifier for the email. */
  id?: string;
  /** The date and time when the email was sent. */
  sentDateTime?: string;
  /** A key indicating the version of the email for tracking changes. */
  changeKey?: string;
  /** The date and time when the email was received. */
  receivedDateTime?: string;
  /** Indicates whether the information is outdated or not. */
  isOutOfDate?: boolean;
  flag?: Flag;
  body?: Body;
  /** Indicates the priority level of the email, such as high, normal, or low. */
  importance?: string;
  /** Shows whether the email has been read or is still marked as unread. */
  isRead?: boolean;
  sender?: Sender;
  endDateTime?: EndDateTime;
  /** The classification of the email based on inferred importance. */
  inferenceClassification?: string;
  /** A unique identifier for the email message on the internet. */
  internetMessageId?: string;
  /** Indicates whether the event lasts for the entire day. */
  isAllDay?: boolean;
  /** Indicates the type of response expected from the API. */
  responseType?: string;
  toRecipients?: Array<ToRecipient>;
  /** Indicates whether the action is performed by a delegate. */
  isDelegated?: boolean;
  /** Indicates whether the email contains any attachments. */
  hasAttachments?: boolean;
  /** Indicates if a response is requested from the recipient. */
  responseRequested?: boolean;
  /** Unique identifier for the position of the email in a thread. */
  conversationIndex?: string;
  /** Specifies the type of meeting request, such as new or update. */
  meetingRequestType?: string;
  /** Unique identifier for the email conversation thread. */
  conversationId?: string;
}

/**
 * Get Email List — uipath.connector.uipath-microsoft-outlook365.get-email-list
 *
 * LOOKUP: `parentFolderId` -> MailFolder.id
 *   by: `displayName`
 *   use: lookup(GetEmailList, 'parentFolderId').byDisplayName(value)
 *   also: .byDisplayName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const GetEmailList = descriptor<GetEmailListInputs, GetEmailListOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.get-email-list",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
  lookups: {
    parentFolderId: {
      objectName: "MailFolder",
      path: "/MailFolders",
      by: ["displayName"],
      value: "id",
      aliases: { byDisplayName: "displayName" },
      strategy: "tree",
    },
  },
} as const);

export interface GetEventByIdCalendarEventInputs {
  /** The ID of calendar to fetch event LOOKUP — sends a `CalendarList.ID`, never a name or e-mail: lookup(<Descriptor>, 'calendarID') (helpers in the LOOKUP lines above the descriptor). */
  calendarID?: ConnectorLookupValue<string>;
  /** Timezone for output */
  timezone?: ConnectorValue<string>;
  /** The event ID to fetch */
  id: ConnectorValue<string>;
}

export interface GetEventByIdCalendarEventOutputs {
  /** The date and time when the calendar entry was last updated. */
  lastModifiedDateTime?: string;
  recurrence?: Recurrence;
  location?: Location;
  locations?: Array<Location>;
  /** Title → e.g. Event title */
  subject?: string;
  /** A URL link to access the calendar event online. */
  webLink?: string;
  /** Indicates if the calendar event is saved as a draft. */
  isDraft?: boolean;
  /** A short preview of the event's description or content. */
  bodyPreview?: string;
  /** The date and time when the calendar event was created. */
  createdDateTime?: string;
  /** Specifies the type of calendar event or action. */
  type?: string;
  /** The time zone in which the event originally ends. */
  originalEndTimeZone?: string;
  /** Indicates if attendees can propose new times for the event. */
  allowNewTimeProposals?: boolean;
  /** The number of minutes before the event to send a reminder. */
  reminderMinutesBeforeStart?: number;
  attendees?: Array<Attendee>;
  /** A unique identifier for the calendar event. */
  id?: string;
  /** The time zone in which the event was originally scheduled. */
  originalStartTimeZone?: string;
  /** Indicates how the event appears on the calendar, like busy or free. */
  showAs?: string;
  responseStatus?: ResponseStatus;
  /** A unique identifier for tracking the calendar action transaction. */
  transactionId?: string;
  start?: Start;
  /** A unique identifier for tracking changes to the event. */
  changeKey?: string;
  /** Indicates whether the user is the organizer of the event. */
  isOrganizer?: boolean;
  /** Indicates the privacy level of the calendar event. */
  sensitivity?: string;
  onlineMeeting?: OnlineMeeting;
  /** Indicates if the meeting is set as an online meeting. Should only be marked as true for Teams meeting. */
  isOnlineMeeting?: boolean;
  body?: Body;
  /** Defines the priority level of the calendar event. */
  importance?: string;
  /** A unique identifier for the calendar event in iCalendar format. */
  iCalUId?: string;
  organizer?: Organizer;
  /** The service used for hosting the online meeting. */
  onlineMeetingProvider?: string;
  end?: End;
  /** Indicates if the event lasts the entire day. */
  isAllDay?: boolean;
  /** Indicates whether the calendar event includes attachments. */
  hasAttachments?: boolean;
  /** Indicates if a response is requested from attendees. */
  responseRequested?: boolean;
  /** Shows whether the event has been cancelled. */
  isCancelled?: boolean;
  /** Indicates if a reminder is set for the event. */
  isReminderOn?: boolean;
  /** Indicates if attendees are hidden from the calendar event. */
  hideAttendees?: boolean;
  /** The web link to join the online meeting associated with the event. */
  onlineMeetingUrl?: string;
}

/**
 * Get Event by ID (Calendar Event) — uipath.connector.uipath-microsoft-outlook365.get-event-by-id-calendar-event-
 *
 * LOOKUP: `calendarID` -> CalendarList.ID
 *   by: `FullName`
 *   use: lookup(GetEventByIdCalendarEvent, 'calendarID').byFullName(value)
 *   also: .byFullName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const GetEventByIdCalendarEvent = descriptor<GetEventByIdCalendarEventInputs, GetEventByIdCalendarEventOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.get-event-by-id-calendar-event-",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
  lookups: {
    calendarID: {
      objectName: "CalendarList",
      path: "/CalendarList",
      by: ["FullName"],
      value: "ID",
      aliases: { byFullName: "FullName" },
      strategy: "tree",
    },
  },
} as const);

export interface GetEventListCalendarEventInputs {
  /** Indicates the date and time as of when to search for events */
  from: ConnectorValue<string>;
  /** Timezone for current event. If left blank will be defaulted to UTC. */
  outputTimezone?: ConnectorValue<string>;
  /** The maximum number of events to be returned. */
  size?: ConnectorValue<string>;
  /** The calendar for the event. If not indicated, the user's default calendar will be used. LOOKUP — sends a `CalendarList.ID`, never a name or e-mail: lookup(<Descriptor>, 'calendarID') (helpers in the LOOKUP lines above the descriptor). */
  calendarID?: ConnectorLookupValue<string>;
  /** Indicates the date and time until which to search for events */
  until: ConnectorValue<string>;
  /** The OData query parameter to retrieve a subset of the CuratedCalendarEvents */
  filter?: ConnectorValue<string>;
}

export interface GetEventListCalendarEventOutputs {
  attendees?: Array<Attendee>;
  locations?: Array<Location>;
  /** The Value uid */
  uid?: string;
  /** The Value subject */
  subject?: string;
  /** The Value web link */
  webLink?: string;
  /** The Value type */
  type?: string;
  /** The Value has attachments */
  hasAttachments?: boolean;
  /** The Value allow new time proposals */
  allowNewTimeProposals?: boolean;
  location?: Location;
  end?: End;
  /** The Value is draft */
  isDraft?: boolean;
  /** The Value last modified date time */
  lastModifiedDateTime?: string;
  /** The Value original start time zone */
  originalStartTimeZone?: string;
  /** The Value is reminder on */
  isReminderOn?: boolean;
  /** The Value original end time zone */
  originalEndTimeZone?: string;
  /** The Value show as */
  showAs?: string;
  /** The Value ID */
  id?: string;
  /** The Value online meeting provider */
  onlineMeetingProvider?: string;
  /** The Value is cancelled */
  isCancelled?: boolean;
  start?: Start;
  body?: Body;
  /** The Value hide attendees */
  hideAttendees?: boolean;
  /** The Value sensitivity */
  sensitivity?: string;
  /** The Value i cal u ID */
  iCalUId?: string;
  /** The Value body preview */
  bodyPreview?: string;
  organizer?: Organizer;
  /** The Value is organizer */
  isOrganizer?: boolean;
  /** The Value is online meeting */
  isOnlineMeeting?: boolean;
  /** The Value importance */
  importance?: string;
  /** The Value series master ID */
  seriesMasterId?: string;
  responseStatus?: ResponseStatus;
  /** The Value created date time */
  createdDateTime?: string;
  /** The Value response requested */
  responseRequested?: boolean;
  /** The Value change key */
  changeKey?: string;
  /** The Value online meeting url */
  onlineMeetingUrl?: string;
  /** The Value occurrence ID */
  occurrenceId?: string;
  /** The Value is all day */
  isAllDay?: boolean;
  /** The Value reminder minutes before start */
  reminderMinutesBeforeStart?: number;
}

/**
 * Get Event List (Calendar Event) — uipath.connector.uipath-microsoft-outlook365.get-event-list-calendar-event-
 *
 * LOOKUP: `calendarID` -> CalendarList.ID
 *   by: `FullName`
 *   use: lookup(GetEventListCalendarEvent, 'calendarID').byFullName(value)
 *   also: .byFullName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const GetEventListCalendarEvent = descriptor<GetEventListCalendarEventInputs, GetEventListCalendarEventOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.get-event-list-calendar-event-",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
  lookups: {
    calendarID: {
      objectName: "CalendarList",
      path: "/CalendarList",
      by: ["FullName"],
      value: "ID",
      aliases: { byFullName: "FullName" },
      strategy: "tree",
    },
  },
} as const);

export interface GetNewestEmailInputs {
  /** The OData query parameter to retrieve an email */
  filter?: ConnectorValue<string>;
  /** Indicates whether to consider only emails with attachments */
  withAttachmentsOnly?: ConnectorValue<boolean>;
  /** Indicates whether to mark the retrieved email as read */
  markAsRead?: ConnectorValue<boolean>;
  /** The folder to get the email from LOOKUP — sends a `MailFolder.id`, never a name or e-mail: lookup(<Descriptor>, 'parentFolderId') (helpers in the LOOKUP lines above the descriptor). */
  parentFolderId: ConnectorLookupValue<string>;
  /** OrderBy */
  orderBy?: ConnectorValue<string>;
  /** Indicates whether to consider unread only */
  unReadOnly?: ConnectorValue<boolean>;
  /** Top */
  top?: ConnectorValue<string>;
  /** The importance of the email */
  importance?: ConnectorValue<string>;
  /** Return only emails in this conversation thread. Use the Conversation Id returned by Reply To Email (or a message trigger) to track incoming replies. */
  conversationId?: ConnectorValue<string>;
}

export interface GetNewestEmailOutputs {
  /** The date and time when the email was last modified. */
  lastModifiedDateTime?: string;
  body?: Body;
  /** The unique identifier of the folder containing the email. */
  parentFolderId?: string;
  /** The subject line of the email message. */
  subject?: string;
  /** The priority level of the email, such as high, normal, or low. */
  importance?: string;
  /** A URL link to view the email in a web browser. */
  webLink?: string;
  /** Indicates whether the email has been read. */
  isRead?: boolean;
  /** Indicates whether the email is a draft. */
  isDraft?: boolean;
  /** The date and time when the email was created. */
  createdDateTime?: string;
  /** A short preview of the email's content. */
  bodyPreview?: string;
  sender?: Sender;
  /** The classification of the email based on its content. */
  inferenceClassification?: string;
  from?: From;
  /** A unique identifier for the email message. */
  internetMessageId?: string;
  toRecipients?: Array<ToRecipient>;
  /** Indicates if a read receipt is requested. */
  isReadReceiptRequested?: boolean;
  /** The unique identifier for the email. */
  id?: string;
  /** Indicates if the email contains attachments. */
  hasAttachments?: boolean;
  /** The date and time when the email was sent. */
  sentDateTime?: string;
  /** A unique identifier for the email's position in a thread. */
  conversationIndex?: string;
  /** A unique identifier for the email conversation thread. */
  conversationId?: string;
  /** A unique identifier for the current version of the email. */
  changeKey?: string;
  /** The date and time when the email was received. */
  receivedDateTime?: string;
  /** Indicates if a delivery receipt was requested for the email. */
  isDeliveryReceiptRequested?: boolean;
  flag?: Flag;
}

/**
 * Get Newest Email — uipath.connector.uipath-microsoft-outlook365.get-newest-email
 *
 * LOOKUP: `parentFolderId` -> MailFolder.id
 *   by: `displayName`
 *   use: lookup(GetNewestEmail, 'parentFolderId').byDisplayName(value)
 *   also: .byDisplayName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const GetNewestEmail = descriptor<GetNewestEmailInputs, GetNewestEmailOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.get-newest-email",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
  lookups: {
    parentFolderId: {
      objectName: "MailFolder",
      path: "/MailFolders",
      by: ["displayName"],
      value: "id",
      aliases: { byDisplayName: "displayName" },
      strategy: "tree",
    },
  },
} as const);

export interface ListEmailAttachmentsInputs {
  /** List of fields to select */
  fields?: ConnectorValue<string>;
  /** The ID of email message */
  id: ConnectorValue<string>;
  /** Indicates whether to exclude the inline attachments, i.e. attachments embedded in the email body */
  excludeInlineAttachments?: ConnectorValue<boolean>;
}

export interface ListEmailAttachmentsOutputs {
  /** The date and time when the attachment was last modified. */
  lastModifiedDateTime?: string;
  /** The size of the attachment in bytes. */
  size?: number;
  /** The unique identifier for the attachment content. */
  contentId?: string;
  /** The name of the attachment file. */
  name?: string;
  /** Indicates if the attachment is displayed within the email body. */
  isInline?: boolean;
  /** Unique identifier for the attachment. */
  id?: string;
  /** The binary content of the attachment encoded in base64. */
  contentBytes?: string;
  /** Specifies the media type of the attachment file. */
  contentType?: string;
}

/** List Email Attachments — uipath.connector.uipath-microsoft-outlook365.list-email-attachments */
export const ListEmailAttachments = descriptor<ListEmailAttachmentsInputs, ListEmailAttachmentsOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.list-email-attachments",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface MarkEmailReadOrUnreadInputs {
  /** Indicates whether the email has been read or not. */
  isRead?: ConnectorValue<string>;
  /** The ID of email to mark as read or unread */
  id: ConnectorValue<string>;
}

export interface MarkEmailReadOrUnreadOutputs {
  /** The date and time when the email was last modified. */
  lastModifiedDateTime?: string;
  body?: Body;
  /** The unique identifier of the folder containing the email. */
  parentFolderId?: string;
  /** The subject line of the email. */
  subject?: string;
  /** The priority level of the email, such as high, normal, or low. */
  importance?: string;
  /** A URL link to view the email in a web browser. */
  webLink?: string;
  /** Indicates whether the email has been read or not. */
  isRead?: string;
  /** Shows if the email is currently saved as a draft. */
  isDraft?: boolean;
  /** The date and time when the email was created. */
  createdDateTime?: string;
  /** A short preview of the email's content. */
  bodyPreview?: string;
  sender?: Sender;
  /** The categories assigned to the email for organization. */
  categories?: Array<string>;
  /** The classification of the email based on inferred importance. */
  inferenceClassification?: string;
  from?: From;
  /** A unique identifier for the email message. */
  internetMessageId?: string;
  toRecipients?: Array<ToRecipient>;
  /** Indicates if a read receipt is requested for the email. */
  isReadReceiptRequested?: boolean;
  replyTo?: Array<ReplyTo>;
  /** A unique identifier for the email message. */
  id?: string;
  /** Indicates whether the email has any attachments. */
  hasAttachments?: boolean;
  /** The date and time when the email was sent. */
  sentDateTime?: string;
  /** A unique index for the email within a conversation. */
  conversationIndex?: string;
  /** A unique identifier for the email conversation thread. */
  conversationId?: string;
  /** A key indicating the version of the email item. */
  changeKey?: string;
  /** The date and time when the email was received. */
  receivedDateTime?: string;
  flag?: Flag;
}

/** Mark Email Read or Unread — uipath.connector.uipath-microsoft-outlook365.mark-email-read-or-unread */
export const MarkEmailReadOrUnread = descriptor<MarkEmailReadOrUnreadInputs, MarkEmailReadOrUnreadOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.mark-email-read-or-unread",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface MoveEmailInputs {
  /** The folder where to move the email LOOKUP — sends a `MailFolder.id`, never a name or e-mail: lookup(<Descriptor>, 'destinationId') (helpers in the LOOKUP lines above the descriptor). */
  destinationId: ConnectorLookupValue<string>;
  /** The email ID to move. This can be retrieved from a trigger or via get newest email activity */
  id: ConnectorValue<string>;
}

export interface MoveEmailOutputs {
  /** The date and time when the email was last modified. */
  lastModifiedDateTime?: string;
  body?: Body;
  /** The unique identifier of the folder containing the email. */
  parentFolderId?: string;
  /** The subject line of the email message. */
  subject?: string;
  /** The priority level of the email, such as high, normal, or low. */
  importance?: string;
  /** A URL link to access the email in a web browser. */
  webLink?: string;
  /** Indicates whether the email has been read. */
  isRead?: boolean;
  /** Shows if the email is currently a draft. */
  isDraft?: boolean;
  /** The date and time when the email was created. */
  createdDateTime?: string;
  /** A short preview of the email's content. */
  bodyPreview?: string;
  sender?: Sender;
  /** The classification of the email based on inferred importance. */
  inferenceClassification?: string;
  from?: From;
  /** Unique identifier for the email message. */
  internetMessageId?: string;
  toRecipients?: Array<ToRecipient>;
  /** Indicates if a read receipt is requested for the email. */
  isReadReceiptRequested?: boolean;
  /** Unique identifier for the email within the system. */
  id?: string;
  /** Indicates if the email contains any attachments. */
  hasAttachments?: boolean;
  /** The date and time when the email was sent. */
  sentDateTime?: string;
  /** A unique identifier for the email's position in a conversation. */
  conversationIndex?: string;
  /** A unique identifier for the email conversation thread. */
  conversationId?: string;
  /** A unique identifier for the version of the email. */
  odataEtag?: string;
  /** The URL that provides context for the email data. */
  odataContext?: string;
  /** A key indicating the current version of the email. */
  changeKey?: string;
  /** The date and time when the email was received. */
  receivedDateTime?: string;
  flag?: Flag;
}

/**
 * Move Email — uipath.connector.uipath-microsoft-outlook365.move-email
 *
 * LOOKUP: `destinationId` -> MailFolder.id
 *   by: `displayName`
 *   use: lookup(MoveEmail, 'destinationId').byDisplayName(value)
 *   also: .byDisplayName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const MoveEmail = descriptor<MoveEmailInputs, MoveEmailOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.move-email",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
  lookups: {
    destinationId: {
      objectName: "MailFolder",
      path: "/MailFolders",
      by: ["displayName"],
      value: "id",
      aliases: { byDisplayName: "displayName" },
      strategy: "tree",
    },
  },
} as const);

export interface ReplyToEmailInputs {
  message?: ConnectorValue<MessageInput2>;
  /** The body of email */
  comment?: ConnectorValue<string>;
  /** The id of the email to reply to */
  id: ConnectorValue<string>;
}

export interface ReplyToEmailOutputs {
  /** The status of the reply operation. */
  status?: string;
  /** A unique identifier for the reply message. */
  id?: string;
  /** Unique identifier for the email conversation thread. Use this with the List Emails / Get Newest Email Conversation ID filter to track incoming replies. */
  conversationId?: string;
  /** A unique identifier for the email message on the internet. */
  internetMessageId?: string;
}

/** Reply to Email — uipath.connector.uipath-microsoft-outlook365.reply-to-email */
export const ReplyToEmail = descriptor<ReplyToEmailInputs, ReplyToEmailOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.reply-to-email",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface RespondToEventInvitationCalendarEventInputs {
  /** Add a comment to your response for the event invitation. */
  comment?: ConnectorValue<string>;
  /** Notify organizer of the response status */
  sendResponse?: ConnectorValue<boolean>;
  /** The new date and time for the start of the event.  If left blank, the existing value will not be updated. */
  startDateTime?: ConnectorValue<string>;
  /** The new date and time for the start of the event.  If left blank, the existing value will not be updated. */
  endDateTime?: ConnectorValue<string>;
  /** The new timezone for the event. If left blank, the existing value will not be updated. */
  timezone?: ConnectorValue<string>;
  /** The response status sent to organizer */
  response: ConnectorValue<string>;
  /** Update all occurance of given event with the same response */
  applyOnSeries?: ConnectorValue<boolean>;
  /** The ID of calendar in which event is present. Should be left empty in case of personal calendar. LOOKUP — sends a `CalendarList.ID`, never a name or e-mail: lookup(<Descriptor>, 'calendarID') (helpers in the LOOKUP lines above the descriptor). */
  calendarID?: ConnectorLookupValue<string>;
  /** The ID of event to be replied */
  id: ConnectorValue<string>;
}

export interface RespondToEventInvitationCalendarEventOutputs {
  /** Notify organizer of the response status */
  sendResponse?: boolean;
}

/**
 * Respond to Event Invitation (Calendar Event) — uipath.connector.uipath-microsoft-outlook365.respond-to-event-invitation-calendar-event-
 *
 * LOOKUP: `calendarID` -> CalendarList.ID
 *   by: `FullName`
 *   use: lookup(RespondToEventInvitationCalendarEvent, 'calendarID').byFullName(value)
 *   also: .byFullName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const RespondToEventInvitationCalendarEvent = descriptor<RespondToEventInvitationCalendarEventInputs, RespondToEventInvitationCalendarEventOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.respond-to-event-invitation-calendar-event-",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
  lookups: {
    calendarID: {
      objectName: "CalendarList",
      path: "/CalendarList",
      by: ["FullName"],
      value: "ID",
      aliases: { byFullName: "FullName" },
      strategy: "tree",
    },
  },
} as const);

export interface SendEmailInputs {
  message: ConnectorValue<MessageInput3>;
  /** Indicates if the email should be saved in the sent items folder. */
  saveToSentItems?: ConnectorValue<boolean>;
}

export interface SendEmailOutputs {
  status?: string;
  /** A unique identifier for the sent email or saved draft. */
  id?: string;
  /** Unique identifier for the email conversation thread. Use this with the List Emails / Get Newest Email Conversation ID filter to track incoming replies. */
  conversationId?: string;
  /** A unique identifier for the email message on the internet. */
  internetMessageId?: string;
}

/** Send Email — uipath.connector.uipath-microsoft-outlook365.send-email */
export const SendEmail = descriptor<SendEmailInputs, SendEmailOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.send-email",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface SetEmailCategoriesInputs {
  /** The categories assigned to the email for organization. */
  categories: ConnectorValue<Array<string>>;
  /** Categories to remove for the email */
  categoriesToRemove?: ConnectorValue<Array<string>>;
  /** Indicates which categories to remove from the email */
  removeCategoriesOption?: ConnectorValue<string>;
  /** The ID of email message to associate with specified categories */
  id: ConnectorValue<string>;
}

export interface SetEmailCategoriesOutputs {
  /** The date and time when the email was last modified. */
  lastModifiedDateTime?: string;
  body?: Body;
  /** The unique identifier of the folder containing the email. */
  parentFolderId?: string;
  /** The subject line of the email message. */
  subject?: string;
  /** The priority level of the email, such as low, normal, or high. */
  importance?: string;
  /** A URL link to access the email in a web browser. */
  webLink?: string;
  /** Indicates whether the email has been read. */
  isRead?: boolean;
  /** Shows if the email is currently a draft. */
  isDraft?: boolean;
  /** The date and time when the email was created. */
  createdDateTime?: string;
  /** A short preview of the email's content. */
  bodyPreview?: string;
  sender?: Sender;
  /** The categories assigned to the email for organization. */
  categories?: Array<string>;
  /** Categories to remove for the email */
  categoriesToRemove?: Array<string>;
  /** The classification of the email based on inferred importance. */
  inferenceClassification?: string;
  from?: From;
  /** Unique identifier for the email message. */
  internetMessageId?: string;
  toRecipients?: Array<ToRecipient>;
  /** Indicates if a read receipt is requested for the email. */
  isReadReceiptRequested?: boolean;
  replyTo?: Array<ReplyTo>;
  /** A unique identifier for the email message. */
  id?: string;
  /** Indicates whether the email has any attachments. */
  hasAttachments?: boolean;
  /** The date and time when the email was sent. */
  sentDateTime?: string;
  /** An index that helps track the email's conversation thread. */
  conversationIndex?: string;
  /** A unique identifier for the email conversation thread. */
  conversationId?: string;
  /** A key indicating the version of the email item. */
  changeKey?: string;
  /** The date and time when the email was received. */
  receivedDateTime?: string;
  flag?: Flag;
}

/** Set Email Categories — uipath.connector.uipath-microsoft-outlook365.set-email-categories */
export const SetEmailCategories = descriptor<SetEmailCategoriesInputs, SetEmailCategoriesOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.set-email-categories",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface TurnOffAutomaticRepliesInputs {
  /** The object body */
  body?: ConnectorValue<BodyInput2>;
}

export interface TurnOffAutomaticRepliesOutputs {
  automaticRepliesSetting?: AutomaticRepliesSetting;
  /** The intended purpose or role of the user in the system. */
  userPurpose?: string;
}

/** Turn off Automatic Replies — uipath.connector.uipath-microsoft-outlook365.turn-off-automatic-replies */
export const TurnOffAutomaticReplies = descriptor<TurnOffAutomaticRepliesInputs, TurnOffAutomaticRepliesOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.turn-off-automatic-replies",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface TurnOnAutomaticRepliesInputs {
  automaticRepliesSetting: ConnectorValue<AutomaticRepliesSettingInput>;
}

export interface TurnOnAutomaticRepliesOutputs {
  automaticRepliesSetting?: AutomaticRepliesSetting;
  /** Indicates the intended use or role of the user. */
  userPurpose?: string;
}

/** Turn on Automatic Replies — uipath.connector.uipath-microsoft-outlook365.turn-on-automatic-replies */
export const TurnOnAutomaticReplies = descriptor<TurnOnAutomaticRepliesInputs, TurnOnAutomaticRepliesOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.turn-on-automatic-replies",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
} as const);

export interface UpdateEventCalendarEventInputs {
  location?: ConnectorValue<LocationInput2>;
  /** Specifies if the meeting should be online (only Microsoft Teams is supported). */
  isOnlineMeeting?: ConnectorValue<boolean>;
  /** Event status displayed in calendar */
  showAs?: ConnectorValue<string>;
  /** The timezone in which the event is scheduled. */
  timezone?: ConnectorValue<string>;
  /** The importance of the event. */
  importance?: ConnectorValue<string>;
  body?: ConnectorValue<BodyInput3>;
  start?: ConnectorValue<StartInput2>;
  /** Specifies if the event lasts the entire day. */
  isAllDay?: ConnectorValue<boolean>;
  /** Update the required attendees via one of the given options */
  changeRequiredAttendees?: ConnectorValue<string>;
  /** Defines the privacy level of the event, such as public or private. */
  sensitivity?: ConnectorValue<string>;
  /** Update the resource attendees via one of the given options */
  changeResourceAttendees?: ConnectorValue<string>;
  end?: ConnectorValue<EndInput2>;
  /** Update the optional attendees via one of the given options */
  changeOptionalAttendees?: ConnectorValue<string>;
  /** The new name of the event. If left blank, the existing value will not be updated. */
  subject?: ConnectorValue<string>;
  /** Updates the categories list with one of the given option */
  changeCategories?: ConnectorValue<string>;
  /** Indicates attendee(s) to be added */
  addRequiredAttendees?: ConnectorValue<string>;
  /** Indicates attendee(s) to be removed */
  removeRequiredAttendees?: ConnectorValue<string>;
  /** Indicates attendee(s) to be added */
  addOptionalAttendees?: ConnectorValue<string>;
  /** Indicates attendee(s) to be removed */
  removeOptionalAttendees?: ConnectorValue<string>;
  /** Indicates attendee(s) to be added */
  addResourceAttendees?: ConnectorValue<string>;
  /** Indicates attendee(s) to be removed */
  removeResourceAttendees?: ConnectorValue<string>;
  /** Indicate the attendee(s) or the attendees list that should replaced */
  requiredAttendees?: ConnectorValue<string>;
  /** Indicate the attendee(s) or the attendees list that should replaced */
  optionalAttendees?: ConnectorValue<string>;
  /** Indicate the attendee(s) or the attendees list that should replaced */
  resourceAttendees?: ConnectorValue<string>;
  /** Indicates the new categories that should replace the existing categories */
  categories?: ConnectorValue<Array<string>>;
  /** Indicates the categories to be added */
  addCategories?: ConnectorValue<Array<string>>;
  /** Indicates categories to be removed */
  removeCategories?: ConnectorValue<Array<string>>;
  /** Attachment */
  file?: ConnectorValue<string>;
  /** The ID of calendar in which respective event is present LOOKUP — sends a `CalendarList.ID`, never a name or e-mail: lookup(<Descriptor>, 'calendarID') (helpers in the LOOKUP lines above the descriptor). */
  calendarID?: ConnectorLookupValue<string>;
  /** Timezone for output event */
  outputTimezone?: ConnectorValue<string>;
  /** The ID of attachment to be removed from event */
  removeAttachmentID?: ConnectorValue<string>;
  /** The ID of event to modify */
  id: ConnectorValue<string>;
  /** Change attachment associated with event */
  changeAttachment?: ConnectorValue<string>;
}

export interface UpdateEventCalendarEventOutputs {
  /** Event status displayed in calendar */
  showAs?: string;
  /** The timezone in which the event is scheduled. */
  timezone?: string;
  /** The importance of the event. */
  importance?: string;
  /** Update the attachment with one of the provided option */
  changeAttachments?: string;
  /** Specifies if the event lasts the entire day. */
  isAllDay?: boolean;
}

/**
 * Update Event (Calendar Event) — uipath.connector.uipath-microsoft-outlook365.update-event-calendar-event-
 *
 * LOOKUP: `calendarID` -> CalendarList.ID
 *   by: `FullName`
 *   use: lookup(UpdateEventCalendarEvent, 'calendarID').byFullName(value)
 *   also: .byFullName  ·  .by(name, value)
 *   strategy: hierarchical walk (the collection is a tree)
 */
export const UpdateEventCalendarEvent = descriptor<UpdateEventCalendarEventInputs, UpdateEventCalendarEventOutputs>()({
  nodeType: "uipath.connector.uipath-microsoft-outlook365.update-event-calendar-event-",
  version: "1.0.0",
  requiresConnection: true,
  requiresFolderKey: true,
  lookups: {
    calendarID: {
      objectName: "CalendarList",
      path: "/CalendarList",
      by: ["FullName"],
      value: "ID",
      aliases: { byFullName: "FullName" },
      strategy: "tree",
    },
  },
} as const);
